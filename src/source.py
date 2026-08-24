"""
src/source.py — 一覧_YYYYMMDD.xlsx を読んで、集計対象の行だけを取り出す

1ファイルを「キー列の値 → Record」の辞書に変える。前日と当日の突き合わせは、
この辞書のキー（キー列の値）が両方にあるかどうかだけで判定する。
"""

import datetime
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from comken.exceptions import ExcelColumnNotFoundError
from comken.toolbox.excel import Excel

from src.settings import ColumnRule, Criteria, SourceLayout

logger = logging.getLogger(__name__)

# 「日」列が文字列で入っていた場合に受け付ける書き方
DATE_TEXT_FORMATS = ("%Y/%m/%d", "%Y-%m-%d", "%Y年%m月%d日", "%Y/%m/%d %H:%M:%S")


@dataclass(frozen=True)
class Record:
    """条件を満たした1行。突合と集計に必要な値だけを持つ。"""

    customer_id: str
    date: datetime.date
    plan_prefix: str  # 集計表の列グループになる（標準 / 上位）
    plan: str  # 種別の原文（明細シート用）
    kind: str


def read_records(path: Path, layout: SourceLayout, criteria: Criteria) -> dict[str, Record]:
    """1ファイルを読み、条件を満たす行を キー列の値 をキーにした辞書で返す。

    Args:
        path: 一覧_YYYYMMDD.xlsx のパス。
        layout: シート名と列名。
        criteria: 対象月・種別・区分の条件。

    Returns:
        {キー列の値: Record}。条件を満たす行がなければ空の辞書。

    Raises:
        ExcelColumnNotFoundError: 必要な列が見出しにない場合。
    """
    records: dict[str, Record] = {}
    duplicate_ids: list[str] = []
    broken_dates = 0
    row_count = 0
    # Excel(read_only=True) は openpyxl の読み取り専用モード相当。read_computed_rows_as_dicts は
    # 見出し行をキーにした dict のリストを返すので、列名でのアクセスは従来どおり
    with Excel(path, read_only=True) as excel:
        rows = _read_dict_rows(excel, layout, criteria)
        for row in rows:
            row_count += 1
            customer_id = _customer_id(row.get(layout.key_column))
            if not customer_id or not _matches(row, layout, criteria):
                continue
            date = _to_date(row.get(layout.date_column))
            if date is None:
                broken_dates += 1
                continue
            # 対象月に入っているかだけを見る。集計表の横軸になるファイル名の日付とは別物
            if (date.year, date.month) != (criteria.target_year, criteria.target_month):
                continue
            if customer_id in records:
                # 同じIDを複数件数えないため、明細・集計には先頭行だけを使う。
                duplicate_ids.append(customer_id)
                continue
            records[customer_id] = Record(
                customer_id=customer_id,
                date=date,
                plan_prefix=_plan_prefix(row.get(layout.plan_column), criteria.plan_prefixes),
                plan=_text(row.get(layout.plan_column)),
                kind=_text(row.get(layout.kind_column)),
            )

    if row_count == 0:
        logger.warning("データ行がありません: %s", path.name)

    if duplicate_ids:
        logger.warning(
            "%s: %s が重複しています（%d件）。先に出てきた行を採用しました",
            path.name,
            layout.key_column,
            len(duplicate_ids),
        )
        logger.debug("重複した%s: %s", layout.key_column, ", ".join(duplicate_ids))
    if broken_dates:
        logger.warning(
            "%s: %s を日付として読めない行が %d 件あり、集計から外しました",
            path.name,
            layout.date_column,
            broken_dates,
        )
    logger.info("%s: 条件に合う行 %d 件", path.name, len(records))
    return records


def _read_dict_rows(
    excel: Excel, layout: SourceLayout, criteria: Criteria
) -> list[dict[str, Any]]:
    """見出しを検証し、1行ずつ辞書化したリストを返す。

    旧 API の ``iter_rows`` ベース実装では 1 行ずつストリーミングしていたが、新 API は
    最初から dict のリストを返す。中身は同じなので、利用側 (``read_records``) は
    ``for row in rows`` の形をそのまま使える。
    """
    raw_rows = excel.read_computed_rows_as_dicts(
        layout.sheet_name, header_row=layout.header_row
    )
    if not raw_rows:
        return []
    # 見出しの前後の空白はここで落とす。Excel の見出しには「備考 」のように空白が紛れ込む
    # ことがあるが、config.ini 側はキー名の空白が落ちるため、空白付きの列名を書く手段がない。
    # 新 API は dict キーに生の見出し名を入れるので、利用側でまとめて剥がす。
    original_keys = list(raw_rows[0].keys())
    stripped_keys = [_text(key) for key in original_keys]
    _validate_columns(stripped_keys, layout, criteria)
    rows: list[dict[str, Any]] = []
    for raw_row in raw_rows:
        # 全部 None の行は Excel の空行（行末の余白など）。集計対象ではないので飛ばす
        if all(value is None for value in raw_row.values()):
            continue
        rows.append(dict(zip(stripped_keys, raw_row.values(), strict=False)))
    return rows


def _validate_columns(headers: list[str], layout: SourceLayout, criteria: Criteria) -> None:
    """必要な列が見出しに揃っているか確かめる。

    あとから足した絞り込みの列も見る。列名を打ち間違えたまま「1件も該当しない」
    「1件も除外されない」と静かに間違うのを防ぐ。
    """
    required = [
        layout.key_column,
        layout.date_column,
        layout.plan_column,
        layout.kind_column,
    ]
    required += [rule.column for rule in criteria.rules]
    missing = [column for column in dict.fromkeys(required) if column not in headers]
    if missing:
        raise ExcelColumnNotFoundError(missing)


def _matches(row: dict, layout: SourceLayout, criteria: Criteria) -> bool:
    """種別と状態、あとから足した条件を満たすか判定する（日付は呼び出し側で見る）。"""
    if _text(row.get(layout.kind_column)) not in criteria.kinds:
        return False
    if _plan_prefix(row.get(layout.plan_column), criteria.plan_prefixes) == "":
        return False
    texts = {rule.column: _text(row.get(rule.column)) for rule in criteria.rules}
    return all(_matches_rule(texts[rule.column], rule) for rule in criteria.rules)


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
