"""
src/report.py — 集計結果を CSV に書き出す

形は「対象月 × 種別 × 判定」を行とし、業務日を列に並べたもの。
対象月が変わっても古い行は残し、その月の案件が無い新しい日付列は空にする。
"""

import datetime
import logging
from pathlib import Path

from comken.toolbox.csv import CSV

from src.diff import ByRow, RowKey

logger = logging.getLogger(__name__)

COL_TARGET_MONTH = "対象月"
COL_PLAN = "種別"
COL_STATUS = "判定"

STATUS_ADDED = "積み上げ"
STATUS_POSTPONED = "延期"


# 既存 CSV から取り込んだ行。対象月が変わってもそのまま残す形
ExistingRow = dict[str, object]


def write_csv(
    path: Path,
    by_row: ByRow,
    dates: list[tuple[datetime.date, bool]],
    row_keys: list[tuple[str, str]],
    target_dates_by_month: dict[str, set[datetime.date]] | None = None,
) -> tuple[int, int]:
    """集計 CSV を1本だけ書く。

    ``dates`` は ``(業務日, 今日比較したか)`` のリスト。1日も飛ばさず並べ、
    今日比較した日かつ、その対象月がその業務日で対象だった場合は
    「ファイルはあったが件数が 0」を ``"0"`` で表す。それ以外（ファイル無し、
    対象月でない業務日、履歴）は空セルにする。
    ``row_keys`` は呼び出し側で組み立てた ``(対象月, 種別)`` の組。
    既存 CSV があれば (対象月, 種別, 判定) をキーにマージし、
    新しい集計に無いキーはそのまま残す（古い対象月の行も消さない）。

    ``target_dates_by_month`` は対象月ごとに「対象月だった業務日」の集合。
    業務日基準の対象月になったので、同じ行でも対象月でない業務日の列は
    空セルにする必要がある。

    戻り値は ``(書き出した行数, 列数)``。途中保存のたびに呼ばれるので、
    「出力しました」のログは呼び出し側で最後の1回に絞って出す（出力ファイル
    1本につき1行だけにするため）。
    """
    columns = [COL_TARGET_MONTH, COL_PLAN, COL_STATUS, *(d.isoformat() for d, _ in dates)]

    # 既存行をキー引きできる形にする
    existing_by_key: dict[tuple[str, str, str], ExistingRow] = {}
    if path.exists():
        for row in read_existing(path):
            key = (str(row[COL_TARGET_MONTH]), str(row[COL_PLAN]), str(row[COL_STATUS]))
            existing_by_key[key] = row

    statuses = (STATUS_ADDED, STATUS_POSTPONED)
    new_rows: list[dict[str, object]] = []
    for month_str, plan in row_keys:
        for status in statuses:
            key = (month_str, plan, status)
            row = existing_by_key.pop(key, {}).copy()
            row[COL_TARGET_MONTH] = month_str
            row[COL_PLAN] = plan
            row[COL_STATUS] = status
            per_date = by_row.get(RowKey(month_str, plan, status), {})
            month_active_dates = (
                target_dates_by_month.get(month_str) if target_dates_by_month else None
            )
            for d, is_compared in dates:
                header = d.isoformat()
                if d in per_date:
                    row[header] = per_date[d]
                elif (
                    is_compared
                    and month_active_dates is not None
                    and d in month_active_dates
                ):
                    # 業務日が比較対象かつ、その対象月がその業務日で対象月だった
                    # → 件数 0 を ``"0"`` で表す（空セルと区別するため）
                    row[header] = "0"
                elif is_compared and header in row:
                    # 業務日は比較対象だが、その対象月がその業務日で対象月でなかった
                    # → 既存セルの値を消す（前回は対象月だった業務日が、対象月の
                    # 集合が変わったことで対象外になったケースを想定）
                    del row[header]
                # else: 今この run では比較していない業務日、または対象月が一度も
                # 対象月になっていない → 既存セルの値は触らない（前回以前の run で
                # 書き込まれた値も、空セルのまま残す）
            new_rows.append(row)
    # 既存行で対象月に含まれないもの（古い対象月ぶん）はそのまま残す
    leftover_rows = list(existing_by_key.values())

    # ``CSV.replace`` は ``Table`` 経由で見出しと行の列集合が一致していないと
    # ``TableRowColumnsError`` を投げる。既存行は CSV にあった列のままで、
    # 今回増えた列は持っていないので、ここで全列を揃える（無い列は空文字）
    rows_for_csv = [
        {column: row.get(column, "") for column in columns}
        for row in new_rows + leftover_rows
    ]
    with CSV(path, columns=columns) as csv:
        csv.replace(rows_for_csv)

    return len(rows_for_csv), len(columns)


def read_existing(path: Path) -> list[ExistingRow]:
    """既存 CSV を読んで行リストを返す。"""
    if not path.exists():
        return []
    with CSV(path) as csv:
        table = csv.read()
    return table.read_rows()


def last_date_in_csv(existing: list[ExistingRow]) -> datetime.date | None:
    """既存 CSV の列のうち**最後の業務日**を ``datetime.date`` で返す。日付列が無ければ None。

    列見出しは ISO 形式のため、``datetime.date.fromisoformat()`` で素直に読む。
    """
    if not existing:
        return None
    first_col_index = 3  # 対象月 / 種別 / 判定 の3列のあと
    dates = [
        date
        for date in (
            _parse_date_header(header)
            for header in list(existing[0].keys())[first_col_index:]
        )
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
