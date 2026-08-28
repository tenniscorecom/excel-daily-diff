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
from comken.core import parse_cell_date
from comken.exceptions import (
    ConfigSectionNotFoundError,
    ExcelColumnNotFoundError,
)
from comken.toolbox.excel import Excel


# 進捗ログ用の ``(現在, 全体)`` の組。素の ``(int, int)`` タプルより名前付きで読みやすい
class Progress(NamedTuple):
    current: int
    total: int

logger = logging.getLogger(__name__)


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
    crew: str  # 集計表のもう1つの列グループになる（作業班。完全一致）


def read_records(
    path: Path,
    plan_prefixes: tuple[str, ...],
    crews: tuple[str, ...],
    kinds: tuple[str, ...],
    rules: tuple[ColumnRule, ...],
    progress: Progress | None = None,
) -> dict[str, Record]:
    """1ファイルを読み、条件を満たす行を 顧客番号 をキーにした辞書で返す。

    対象月の絞り込みはここでは行わない（複数月ぶんを1回の読み込みで
    振り分けるため、呼び出し側で対象月を見て分ける）。
    ``progress`` を渡すと、ログに進捗 ``(n/total)`` を前置する。
    """
    # 設定は起動時に決まる値で、行ごとに変わらない。ループの外で1回だけ読む
    raw_sheet_name = config.SOURCE.SHEET_NAME
    sheet_name = raw_sheet_name if isinstance(raw_sheet_name, str) else raw_sheet_name[0]
    header_row = config.SOURCE.HEADER_ROW
    key_column = config.SOURCE.KEY_COLUMN
    date_column = config.SOURCE.DATE_COLUMN
    plan_column = config.SOURCE.PLAN_COLUMN
    crew_column = config.SOURCE.CREW_COLUMN
    kind_column = config.SOURCE.KIND_COLUMN
    required_columns = (key_column, date_column, plan_column, crew_column, kind_column)

    records: dict[str, Record] = {}
    duplicate_ids: list[str] = []
    broken_dates = 0
    row_count = 0
    with Excel(path, read_only=True) as excel:
        rows = _read_dict_rows(
            excel, sheet_name, rules, header_row, required_columns
        )
        for row in rows:
            row_count += 1
            customer_id = _customer_id(row.get(key_column))
            if not customer_id:
                continue
            date = parse_cell_date(row.get(date_column))
            if date is None:
                broken_dates += 1
                continue
            plan_prefix = _plan_prefix(row.get(plan_column), plan_prefixes)
            crew = _crew_match(row.get(crew_column), crews)
            kind_text = _text(row.get(kind_column))
            if plan_prefix == "" or crew == "" or kind_text not in kinds:
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
                crew=crew,
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
    progress_prefix = (
        f"({progress[0]}/{progress[1]}) " if progress is not None else ""
    )
    logger.info(
        "%s%s[%s]: 条件に合う行 %d 件",
        progress_prefix,
        path.name,
        sheet_name,
        len(records),
    )
    return records


def load_rules() -> tuple[ColumnRule, ...]:
    """あとから足した絞り込みを、config.ini の4セクションから集める。

    4セクション = INCLUDE_CONTAINS / EXCLUDE_CONTAINS / INCLUDE_EXACT / EXCLUDE_EXACT。
    セクションが無ければ空タプルを返す。
    """
    rules: list[ColumnRule] = []
    for section, is_contains, is_exclude in (
        ("INCLUDE_CONTAINS", True, False),
        ("EXCLUDE_CONTAINS", True, True),
        ("INCLUDE_EXACT", False, False),
        ("EXCLUDE_EXACT", False, True),
    ):
        try:
            values = vars(getattr(config, section))
        except ConfigSectionNotFoundError:
            continue
        for column, value in values.items():
            if column.startswith("_"):
                continue
            if isinstance(value, list):
                words = tuple(
                    str(item).strip() for item in value if str(item).strip()
                )
            else:
                words = tuple(
                    word.strip()
                    for word in str(value).split(",")
                    if word.strip()
                )
            if words:
                rules.append(ColumnRule(column, words, is_contains, is_exclude))
    return tuple(rules)


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
    ``header_row`` と ``required_columns`` は ``read_records`` がループの外で
    1回だけ読んで渡したもの（行ごとに変わらない）。
    """
    raw_rows = excel.read(sheet_name, header_row=header_row).read_rows()
    if not raw_rows:
        return []
    original_keys = list(raw_rows[0].keys())
    stripped_keys = [_text(key) for key in original_keys]
    # 必要な列（基本4列 + あとから足した絞り込みの列）が見出しに揃っているか
    # 確かめる。打ち間違えたまま「1件も該当しない」「1件も除外されない」と
    # 静かに間違うのを防ぐ。
    required = [*required_columns, *(rule.column for rule in rules)]
    missing = [
        column for column in dict.fromkeys(required) if column not in stripped_keys
    ]
    if missing:
        raise ExcelColumnNotFoundError(missing)
    rows: list[dict[str, Any]] = []
    for raw_row in raw_rows:
        if all(value is None for value in raw_row.values()):
            continue
        rows.append(dict(zip(stripped_keys, raw_row.values(), strict=False)))
    return rows


def _matches_rules(row: dict, rules: tuple[ColumnRule, ...]) -> bool:
    """あとから足した条件をすべて満たすか判定する。"""
    for rule in rules:
        text = _text(row.get(rule.column))
        if rule.is_contains:
            is_hit = any(word in text for word in rule.words)
        else:
            is_hit = text in rule.words
        if rule.is_exclude:
            is_hit = not is_hit
        if not is_hit:
            return False
    return True


def _plan_prefix(value: object, prefixes: tuple[str, ...]) -> str:
    """種別がどの語で始まるかを返す。どれにも当てはまらなければ空文字。"""
    plan = _text(value)
    for prefix in prefixes:
        if plan.startswith(prefix):
            return prefix
    return ""


def _crew_match(value: object, crews: tuple[str, ...]) -> str:
    """作業班が対象語のいずれかと完全一致するかを返す。どれにも当てはまらなければ空文字。

    ``_plan_prefix`` は前方一致で「どのグループPrefixに当てはまるか」を返すが、
    作業班は完全一致で「対象語そのもの」を返す（部分一致や「◯◯作業班A」「◯◯作業班B」
    のような派生語を別グループにしないため）。
    """
    crew = _text(value)
    return crew if crew in crews else ""


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
