"""
src/report.py — 集計結果を CSV に書き出す

形は「対象月 × 種別 × 判定」を行とし、日付を列に並べたもの。
対象月が変わっても古い行は残し、その月の案件が無い新しい日付列は空にする。
"""

import datetime
import logging
from pathlib import Path

from comken.toolbox.csv import CSV

logger = logging.getLogger(__name__)

COL_TARGET_MONTH = "対象月"
COL_PLAN = "種別"
COL_STATUS = "判定"

STATUS_ADDED = "積み上げ"
STATUS_POSTPONED = "延期"


def _date_header(date: datetime.date) -> str:
    """日付を列見出し用の文字列にする（ISO 形式 ``2026-04-21``）。"""
    return date.isoformat()


# 既存 CSV から取り込んだ行。対象月が変わってもそのまま残す形
ExistingRow = dict[str, object]


def write_csv(
    path: Path,
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]],
    dates: list[tuple[datetime.date, bool]],
    row_keys: list[tuple[str, str]],
) -> None:
    """集計 CSV を1本だけ書く。

    ``dates`` は ``(日付, 今日比較したか)`` のリスト。1日も飛ばさず並べ、
    今日比較した日は「ファイルはあったが件数が 0」なら ``"0"`` を入れ、
    それ以外（ファイル無し・履歴）は空セルにする（README の仕様）。
    ``row_keys`` は呼び出し側で組み立てた ``(対象月, 種別)`` の組。
    既存 CSV があれば (対象月, 種別, 判定) をキーにマージし、
    新しい集計に無いキーはそのまま残す（古い対象月の行も消さない）。
    """
    columns = [COL_TARGET_MONTH, COL_PLAN, COL_STATUS, *(_date_header(d) for d, _ in dates)]

    # 既存行をキー引きできる形にする
    existing_by_key: dict[tuple[str, str, str], ExistingRow] = {}
    if path.exists():
        existing_by_key = _read_existing_by_key(path)

    statuses = (STATUS_ADDED, STATUS_POSTPONED)
    new_rows: list[dict[str, object]] = []
    leftover_rows: list[dict[str, object]] = []
    for month_str, plan in row_keys:
        for status in statuses:
            key = (month_str, plan, status)
            row = existing_by_key.pop(key, {}).copy()
            row[COL_TARGET_MONTH] = month_str
            row[COL_PLAN] = plan
            row[COL_STATUS] = status
            month = _month_tuple_from_str(month_str)
            per_date = by_row.get((month, plan, status), {})
            for d, is_compared in dates:
                header = _date_header(d)
                if d in per_date:
                    row[header] = per_date[d]
                elif is_compared:
                    # ファイルはあったが件数が 0 の日 → "0"
                    row[header] = "0"
                # ファイルが無い日・履歴は空のまま（既定値 ""）
            new_rows.append(row)
    # 既存行で対象月に含まれないもの（古い対象月ぶん）はそのまま残す
    leftover_rows = list(existing_by_key.values())

    with CSV(path, columns=columns) as csv:
        csv.replace(_materialize(new_rows + leftover_rows, columns))

    row_count = len(new_rows) + len(leftover_rows)
    logger.info("出力しました: %s（%d 行 × %d 列）", path, row_count, len(columns))


def _month_tuple_from_str(month_str: str) -> tuple[int, int]:
    """``YYYY-MM`` 形式の文字列を ``(year, month)`` のタプルにする。"""
    year_str, month_str_only = month_str.split("-")
    return (int(year_str), int(month_str_only))


def _materialize(
    rows: list[dict[str, object]], columns: list[str]
) -> list[dict[str, object]]:
    """``CSV.replace`` が要求する「列が全部揃った辞書」のリストに整える。"""
    materialized: list[dict[str, object]] = []
    for row in rows:
        materialized.append({column: row.get(column, "") for column in columns})
    return materialized


def _read_existing_by_key(path: Path) -> dict[tuple[str, str, str], ExistingRow]:
    """既存 CSV を ``(対象月, 種別, 判定) → 行`` の辞書で返す。"""
    existing_by_key: dict[tuple[str, str, str], ExistingRow] = {}
    for row in read_existing(path):
        key = (str(row[COL_TARGET_MONTH]), str(row[COL_PLAN]), str(row[COL_STATUS]))
        existing_by_key[key] = row
    return existing_by_key


def read_existing(path: Path) -> list[ExistingRow]:
    """既存 CSV を読んで行リストを返す。"""
    if not path.exists():
        return []
    with CSV(path) as csv:
        table = csv.read()
    return [dict(row) for row in table.read()]


def last_date_in_csv(existing: list[ExistingRow]) -> datetime.date | None:
    """既存 CSV の列のうち**最後の日付**を ``datetime.date`` で返す。日付列が無ければ None。

    列見出しは ISO 形式のため、``datetime.date.fromisoformat()`` で素直に読む。
    """
    if not existing:
        return None
    first_col_index = 3  # 対象月 / 種別 / 判定 の3列のあと
    headers = list(existing[0].keys())[first_col_index:]
    dates = [
        date
        for date in (_parse_date_header(header) for header in headers)
        if date is not None
    ]
    return max(dates) if dates else None


def _parse_date_header(header: str) -> datetime.date | None:
    """``YYYY-MM-DD`` 形式の見出しを date にする。解釈できないものは None。"""
    text = header.strip()
    if not text:
        return None
    try:
        return datetime.date.fromisoformat(text)
    except ValueError:
        return None