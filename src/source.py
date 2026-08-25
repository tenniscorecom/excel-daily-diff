"""
src/source.py — 一覧_YYYYMMDD.xlsx を読んで、集計対象の行だけを取り出す

1ファイルを「顧客番号 → Record」の辞書に変える。前日と当日の突き合わせは、
この辞書のキー（顧客番号）が両方にあるかどうかだけで判定する。
"""

import datetime
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, NamedTuple

from comken import config
from comken.exceptions import (
    ConfigSectionNotFoundError,
    ExcelColumnNotFoundError,
    SheetNotFoundError,
)
from comken.toolbox.excel import Excel

logger = logging.getLogger(__name__)

# 「日」列が文字列で入っていた場合に受け付ける書き方
DATE_TEXT_FORMATS = ("%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日", "%Y/%m/%d %H:%M:%S")


class ColumnRule(NamedTuple):
    """1つの列に対する、あとから足した絞り込み条件。

    例:「地域」に「離島」が含まれる行を外す → column="地域", words=("離島",),
    is_contains=True, is_exclude=True
    """

    column: str
    words: tuple[str, ...]
    is_contains: bool  # True=語を含むか / False=語と完全に一致するか
    is_exclude: bool  # True=当てはまる行を外す / False=当てはまる行だけ残す


@dataclass(frozen=True)
class Record:
    """条件を満たした1行。突合と集計に必要な値だけを持つ。"""

    customer_id: str
    date: datetime.date  # 案件そのものの日付。対象月の判定に使う
    plan_prefix: str  # 集計表の列グループになる（標準 / 上位）


def read_records(
    path: Path,
    plan_prefixes: tuple[str, ...],
    kinds: tuple[str, ...],
    rules: tuple[ColumnRule, ...],
    progress: tuple[int, int] | None = None,
) -> dict[str, Record]:
    """1ファイルを読み、条件を満たす行を 顧客番号 をキーにした辞書で返す。

    対象月の絞り込みはここでは行わない（複数月ぶんを1回の読み込みで
    振り分けるため、呼び出し側で対象月を見て分ける）。
    ``progress`` を渡すと、ログに進捗 ``(n/total)`` を前置する。
    """
    # 設定は起動時に決まる値で、行ごとに変わらない。ループの外で1回だけ読む。
    sheet_names = _source_sheet_names()
    header_row = _source_header_row()
    key_column = _source_key_column()
    date_column = _source_date_column()
    plan_column = _source_plan_column()
    kind_column = _source_kind_column()
    required_columns = (key_column, date_column, plan_column, kind_column)

    records: dict[str, Record] = {}
    duplicate_ids: list[str] = []
    broken_dates = 0
    row_count = 0
    with Excel(path, read_only=True) as excel:
        sheet_name = _select_sheet_name(excel, sheet_names)
        rows = _read_dict_rows(excel, sheet_name, rules, header_row, required_columns)
        for row in rows:
            row_count += 1
            customer_id = _customer_id(row.get(key_column))
            if not customer_id:
                continue
            date = _to_date(row.get(date_column))
            if date is None:
                broken_dates += 1
                continue
            plan_prefix = _plan_prefix(row.get(plan_column), plan_prefixes)
            kind_text = _text(row.get(kind_column))
            if plan_prefix == "" or kind_text not in kinds:
                continue
            if not _matches_rules(row, rules):
                continue
            if customer_id in records:
                duplicate_ids.append(customer_id)
                continue
            records[customer_id] = Record(
                customer_id=customer_id,
                date=date,
                plan_prefix=plan_prefix,
            )

    if row_count == 0:
        logger.warning("データ行がありません: %s", path.name)

    if duplicate_ids:
        logger.warning(
            "%s: %s が重複しています（%d件）。先に出てきた行を採用しました",
            path.name,
            key_column,
            len(duplicate_ids),
        )
        logger.debug("重複した%s: %s", key_column, ", ".join(duplicate_ids))
    if broken_dates:
        logger.warning(
            "%s: %s を日付として読めない行が %d 件あり、集計から外しました",
            path.name,
            date_column,
            broken_dates,
        )
    logger.info(
        "%s%s[%s]: 条件に合う行 %d 件",
        _progress_prefix(progress),
        path.name,
        sheet_name,
        len(records),
    )
    return records


def _progress_prefix(progress: tuple[int, int] | None) -> str:
    """``read_records`` のログに進捗 ``"(3/42) "`` を前置する。``None`` なら空文字。"""
    if progress is None:
        return ""
    current, total = progress
    return f"({current}/{total}) "


def load_rules() -> tuple[ColumnRule, ...]:
    """あとから足した絞り込みを、config.ini の4セクションから集める。

    4セクション = INCLUDE_CONTAINS / EXCLUDE_CONTAINS / INCLUDE_EXACT / EXCLUDE_EXACT。
    セクションが無ければ空タプルを返す。
    """
    rules: list[ColumnRule] = []
    rules += _rules_in("INCLUDE_CONTAINS", is_contains=True, is_exclude=False)
    rules += _rules_in("EXCLUDE_CONTAINS", is_contains=True, is_exclude=True)
    rules += _rules_in("INCLUDE_EXACT", is_contains=False, is_exclude=False)
    rules += _rules_in("EXCLUDE_EXACT", is_contains=False, is_exclude=True)
    return tuple(rules)


def _rules_in(section: str, is_contains: bool, is_exclude: bool) -> list[ColumnRule]:
    try:
        values = vars(getattr(config, section))
    except ConfigSectionNotFoundError:
        return []
    rules = []
    for column, value in values.items():
        if column.startswith("_"):
            continue
        words = _to_words(value)
        if words:
            rules.append(ColumnRule(column, words, is_contains, is_exclude))
    return rules


def _to_words(value: object) -> tuple[str, ...]:
    items = value if isinstance(value, list) else str(value).split(",")
    return tuple(str(item).strip() for item in items if str(item).strip())


def _source_sheet_names() -> tuple[str, ...]:
    """``[SOURCE] SHEET_NAME`` の候補一覧を config に書いた順に返す。

    ``SHEET_NAME = Sheet1`` のように1つだけ書いた古い形式もそのまま動くように、
    文字列で書かれた場合は1要素のタプルに変換する。
    ``[a, b]`` のようにリストで書いた場合は順序を保ったまま要素を返す。
    """
    value = config.SOURCE.SHEET_NAME
    if isinstance(value, list):
        return tuple(str(item).strip() for item in value if str(item).strip())
    return (str(value).strip(),)


def _select_sheet_name(excel: Excel, candidates: tuple[str, ...]) -> str:
    """候補を上から順に試し、最初に見つかったシート名を返す。

    候補が全部見つからないときは、最後の ``SheetNotFoundError`` をそのまま送出する
    （メッセージに実在するシート名の一覧が含まれるので、利用者が config を直せる）。
    候補が空のときは、ブックに存在するシートを一覧にした ``SheetNotFoundError`` を投げる。
    """
    last_error: SheetNotFoundError | None = None
    for name in candidates:
        try:
            excel.sheet(name)
            return name
        except SheetNotFoundError as error:
            last_error = error
    if last_error is not None:
        raise last_error
    # 候補が空（config の書き方が悪い）のときは、実在するシートを一覧にした例外で知らせる
    raise SheetNotFoundError("", excel._workbook.sheetnames)  # noqa: SLF001


def _source_header_row() -> int:
    return int(config.SOURCE.HEADER_ROW)


def _source_key_column() -> str:
    return str(config.SOURCE.KEY_COLUMN)


def _source_date_column() -> str:
    return str(config.SOURCE.DATE_COLUMN)


def _source_plan_column() -> str:
    return str(config.SOURCE.PLAN_COLUMN)


def _source_kind_column() -> str:
    return str(config.SOURCE.KIND_COLUMN)


def _read_dict_rows(
    excel: Excel,
    sheet_name: str,
    rules: tuple[ColumnRule, ...],
    header_row: int,
    required_columns: tuple[str, ...],
) -> list[dict[str, Any]]:
    """見出しを検証し、1行ずつ辞書化したリストを返す。

    Excel の見出しには「備考 」のように空白が紛れ込むことがあるため、
    読み込み側で見出しの前後の空白を落とす（config.ini 側はキー名の空白が落ちる）。
    ``sheet_name`` は ``_select_sheet_name`` で確定済みのものを使う（候補の上から
    試して見つかったもの）。``header_row`` と ``required_columns`` は
    ``read_records`` がループの外で1回だけ読んで渡したもの（行ごとに変わらない）。
    """
    raw_rows = excel.read_computed_rows_as_dicts(sheet_name, header_row=header_row)
    if not raw_rows:
        return []
    original_keys = list(raw_rows[0].keys())
    stripped_keys = [_text(key) for key in original_keys]
    _validate_columns(stripped_keys, rules, required_columns)
    rows: list[dict[str, Any]] = []
    for raw_row in raw_rows:
        if all(value is None for value in raw_row.values()):
            continue
        rows.append(dict(zip(stripped_keys, raw_row.values(), strict=False)))
    return rows


def _validate_columns(
    headers: list[str],
    rules: tuple[ColumnRule, ...],
    required_columns: tuple[str, ...],
) -> None:
    """必要な列が見出しに揃っているか確かめる。

    あとから足した絞り込みの列も見る。列名を打ち間違えたまま「1件も該当しない」
    「1件も除外されない」と静かに間違うのを防ぐ。
    ``required_columns`` は ``read_records`` がループの外で読んだものをそのまま渡す。
    """
    required = [*required_columns, *(rule.column for rule in rules)]
    missing = [column for column in dict.fromkeys(required) if column not in headers]
    if missing:
        raise ExcelColumnNotFoundError(missing)


def _matches_rules(row: dict, rules: tuple[ColumnRule, ...]) -> bool:
    """あとから足した条件をすべて満たすか判定する。"""
    for rule in rules:
        text = _text(row.get(rule.column))
        if not _matches_rule(text, rule):
            return False
    return True


def _matches_rule(text: str, rule: ColumnRule) -> bool:
    """あとから足した条件1つを満たすか判定する。"""
    if rule.is_contains:
        is_hit = any(word in text for word in rule.words)
    else:
        is_hit = text in rule.words
    return not is_hit if rule.is_exclude else is_hit


def _plan_prefix(value: object, prefixes: tuple[str, ...]) -> str:
    """種別がどの語で始まるかを返す。どれにも当てはまらなければ空文字。"""
    plan = _text(value)
    for prefix in prefixes:
        if plan.startswith(prefix):
            return prefix
    return ""


def _to_date(value: object) -> datetime.date | None:
    """セルの値を日付にする。空セルや日付として読めない値は None。"""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = _text(value)
    if not text:
        return None
    for date_format in DATE_TEXT_FORMATS:
        try:
            return datetime.datetime.strptime(text, date_format).date()  # noqa: DTZ007
        except ValueError:
            continue
    return None


def _text(value: object) -> str:
    """セルの値を、前後の空白を落とした文字列にする。空セルは空文字。"""
    if value is None:
        return ""
    return str(value).strip()


def _customer_id(value: object) -> str:
    """Excelで整数が浮動小数になっても同じキー列の値として扱う。"""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return _text(value)