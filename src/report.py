"""
src/report.py — 集計結果を CSV に書き出す

形は「当月/来月ラベル × 種別 × 判定」を行とし、業務日を列に並べたもの。
行の数は **当月/来月 × 種別数 × 2判定**で決まる（``[FILTER] PLAN_PREFIXES``
が既定の 2 件のとき 8 行。``PLAN_PREFIXES`` の件数が増減すれば行数も変わる）。
"""

import datetime
import logging
from pathlib import Path

from comken.toolbox.csv import CSV

from src.diff import LABEL_CURRENT_MONTH, LABEL_NEXT_MONTH, ByRow, RowKey

logger = logging.getLogger(__name__)

# 出力CSV の行ラベル（順序込み）。 ``run.py`` でこの順に行が並ぶ。
ROW_LABELS: tuple[str, ...] = (LABEL_CURRENT_MONTH, LABEL_NEXT_MONTH)

# 列見出し。「対象月」「種別」「判定」の3列構成。
COL_TARGET_MONTH = "対象月"
COL_PLAN = "種別"
COL_STATUS = "判定"

STATUS_ADDED = "積み上げ"
STATUS_POSTPONED = "延期"


# 既存 CSV から取り込んだ行。行ラベルが変わったら（過去バージョンの CSV を
# 持ち越したときなど）は無視される（後述）。
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
    今日比較した日かつ、そのラベルがその業務日で対象だった場合は
    「ファイルはあったが件数が 0」を ``"0"`` で表す。それ以外（ファイル無し、
    ラベルでない業務日、履歴）は空セルにする。

    ``row_keys`` は ``(対象月ラベル, 種別)`` の組のリスト（順序は呼び出し側で
    決めている = 当月/来月の次に ``PLAN_PREFIXES`` の順）。書き出される行は
    ``row_keys`` のそれぞれに {積み上げ, 延期} の 2 行を足した分になる
    （例: ``PLAN_PREFIXES = [標準, 上位]`` のとき 4 組 × 2 = 8 行）。

    ``target_dates_by_month`` はラベルごとに「そのラベルが対象だった業務日」
    の集合。同じ行でも、ラベルでない業務日の列は空セルにする必要がある。

    戻り値は ``(書き出した行数, 列数)``。途中保存のたびに呼ばれるので、
    「出力しました」のログは呼び出し側で最後の1回に絞って出す（出力ファイル
    1本につき1行だけにするため）。

    過去バージョン（``種別``列なしの ``集計.csv``、もしくは別構成の CSV）を
    引き継いだ場合、``row_keys`` に無い行はそのまま残らない（黙って消える）。
    列構成を大きく変えた以上、``集計.csv`` を消してから作り直す運用が
    推奨（``config.ini.example`` の注意書き参照）。
    """
    columns = [
        COL_TARGET_MONTH,
        COL_PLAN,
        COL_STATUS,
        *(d.isoformat() for d, _ in dates),
    ]

    # 既存行を ``(対象月ラベル, 種別, 判定)`` で引ける形にする
    existing_by_key: dict[tuple[str, str, str], ExistingRow] = {}
    if path.exists():
        for row in read_existing(path):
            key = (
                str(row[COL_TARGET_MONTH]),
                str(row[COL_PLAN]),
                str(row[COL_STATUS]),
            )
            existing_by_key[key] = row

    statuses = (STATUS_ADDED, STATUS_POSTPONED)
    new_rows: list[dict[str, object]] = []
    for label, plan in row_keys:
        for status in statuses:
            key = (label, plan, status)
            row = existing_by_key.pop(key, {}).copy()
            row[COL_TARGET_MONTH] = label
            row[COL_PLAN] = plan
            row[COL_STATUS] = status
            per_date = by_row.get(RowKey(label, plan, status), {})
            label_active_dates = (
                target_dates_by_month.get(label) if target_dates_by_month else None
            )
            for d, is_compared in dates:
                header = d.isoformat()
                if d in per_date:
                    row[header] = per_date[d]
                elif (
                    is_compared
                    and label_active_dates is not None
                    and d in label_active_dates
                ):
                    # 業務日が比較対象かつ、そのラベルがその業務日で対象だった
                    # → 件数 0 を ``"0"`` で表す（空セルと区別するため）
                    row[header] = "0"
                elif is_compared and header in row:
                    # 業務日は比較対象だが、そのラベルがその業務日で対象でなかった
                    # → 既存セルの値を消す（前回は対象だった業務日が、対象月の
                    # 集合が変わったことで対象外になったケースを想定）
                    del row[header]
                # else: 今この run では比較していない業務日、またはそのラベルが
                # 一度も対象になっていない → 既存セルの値は触らない（前回以前の
                # run で書き込まれた値も、空セルのまま残す）
            new_rows.append(row)
    # 行は常に固定（``row_keys`` × {積み上げ, 延期}）で、``row_keys`` に無い
    # 既存行を残す必要は無くなった（列構成が変わった以上、過去形式の行は
    # 引き継がない）。``existing_by_key`` に残った分は過去バージョンの CSV を
    # 持ち越した場合だけで、黙って捨てる。

    # ``CSV.replace`` は ``Table`` 経由で見出しと行の列集合が一致していないと
    # ``TableRowColumnsError`` を投げる。既存行は CSV にあった列のままで、
    # 今回増えた列は持っていないので、ここで全列を揃える（無い列は空文字）
    rows_for_csv = [
        {column: row.get(column, "") for column in columns}
        for row in new_rows
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
