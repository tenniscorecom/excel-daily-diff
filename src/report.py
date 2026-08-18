"""
src/report.py — 集計結果を Excel に書き出す

集計シートは「行＝積み上げ／延期、列＝種別ごとの日付」の形にする。
列の日付は**ファイル名の日付**（いつ時点の一覧か）であって、案件の日付ではない。
なぜ何件なのかを後から追えるように、明細シートに元の行を残す。
"""

import datetime
import logging
from pathlib import Path

from comken.toolbox.excel import ExcelWriter, Sheet

from src.diff import DailyDiff
from src.settings import Criteria, Settings, SourceLayout
from src.source import Record

logger = logging.getLogger(__name__)

SUMMARY_SHEET = "集計"
DETAIL_SHEET = "明細"

STATUS_ADDED = "積み上げ"
STATUS_POSTPONED = "延期"

# 集計シートの位置（1始まり）
PLAN_ROW = 1
DATE_ROW = 2
ADDED_ROW = 3
POSTPONED_ROW = 4
LABEL_COL = 1
FIRST_DATE_COL = 2

DATE_NUMBER_FORMAT = "m/d"
COMPARED_DATE_HEADER = "比較日"  # 明細シートで、このツールが付ける列
STATUS_HEADER = "判定"
FROZEN_ROWS = 2  # 集計シートの見出し（種別行・日付行）を固定する

Counts = dict[tuple[str, datetime.date], int]


def write_report(path: Path, diffs: list[DailyDiff], settings: Settings) -> None:
    """集計シートと明細シートを持つブックを作る。

    Args:
        path: 出力先の xlsx パス。
        diffs: 日ごとの差分（古い順）。
        settings: 横軸の期間・種別の並び順・明細シートの見出しに使う列名。
    """
    dates = date_range(settings.start_date, settings.end_date)
    with ExcelWriter.create(path, sheet_name=SUMMARY_SHEET) as f:
        _write_summary(f.sheet(SUMMARY_SHEET), diffs, settings.criteria, dates)
        _write_detail(f.add_sheet(DETAIL_SHEET), diffs, settings.layout)
        f.save()
    logger.info("出力しました: %s", path)


def date_range(start: datetime.date, end: datetime.date) -> list[datetime.date]:
    """開始日から終了日までの日付を1日ずつ並べて返す。"""
    days = (end - start).days + 1
    return [start + datetime.timedelta(days=i) for i in range(days)]


def detail_headers(layout: SourceLayout) -> list[str]:
    """明細シートの見出し。読み取り元の列名をそのまま使う（config.ini を変えれば追随する）。"""
    return [
        COMPARED_DATE_HEADER,
        STATUS_HEADER,
        layout.key_column,
        layout.date_column,
        layout.plan_column,
        layout.kind_column,
    ]


def _write_summary(
    sheet: Sheet, diffs: list[DailyDiff], criteria: Criteria, dates: list[datetime.date]
) -> None:
    """種別ごとに日付を並べた集計表を書く。"""
    added_counts, postponed_counts = _count_by_plan_and_date(diffs)
    # 比較できなかった日（ファイルが無い日）は 0 と区別できるよう空セルのままにする
    compared_dates = {diff.date for diff in diffs}

    sheet.write_cell(ADDED_ROW, LABEL_COL, STATUS_ADDED)
    sheet.write_cell(POSTPONED_ROW, LABEL_COL, STATUS_POSTPONED)

    col = FIRST_DATE_COL
    for plan_prefix in criteria.plan_prefixes:
        # 種別名はグループの先頭列にだけ置く（セルを結合すると並べ替え・集計がしにくい）
        sheet.write_cell(PLAN_ROW, col, plan_prefix)
        sheet.set_bold(PLAN_ROW, col)
        for date in dates:
            sheet.write_cell(DATE_ROW, col, date)
            sheet.set_number_format(DATE_ROW, col, DATE_NUMBER_FORMAT)
            if date in compared_dates:
                sheet.write_cell(ADDED_ROW, col, added_counts.get((plan_prefix, date), 0))
                sheet.write_cell(POSTPONED_ROW, col, postponed_counts.get((plan_prefix, date), 0))
            col += 1

    sheet.freeze_header(FROZEN_ROWS)


def _write_detail(sheet: Sheet, diffs: list[DailyDiff], layout: SourceLayout) -> None:
    """どの行が積み上げ・延期になったかの一覧を書く。"""
    headers = detail_headers(layout)
    rows = []
    for diff in diffs:
        rows += [_detail_row(r, diff.date, STATUS_ADDED, layout) for r in _sorted(diff.added)]
        rows += [
            _detail_row(r, diff.date, STATUS_POSTPONED, layout) for r in _sorted(diff.postponed)
        ]
    if not rows:
        sheet.write_row(1, headers)
        return
    sheet.write_table(rows, headers=headers)
    sheet.auto_width()
    sheet.freeze_header()


def _count_by_plan_and_date(diffs: list[DailyDiff]) -> tuple[Counts, Counts]:
    """(積み上げ, 延期) の件数を、種別と日付の組ごとに数える。"""
    added: Counts = {}
    postponed: Counts = {}
    for diff in diffs:
        for record in diff.added:
            key = (record.plan_prefix, diff.date)
            added[key] = added.get(key, 0) + 1
        for record in diff.postponed:
            key = (record.plan_prefix, diff.date)
            postponed[key] = postponed.get(key, 0) + 1
    return added, postponed


def _detail_row(
    record: Record, compared_date: datetime.date, status: str, layout: SourceLayout
) -> dict[str, object]:
    return {
        COMPARED_DATE_HEADER: compared_date,
        STATUS_HEADER: status,
        layout.key_column: record.customer_id,
        layout.date_column: record.date,
        layout.plan_column: record.plan,
        layout.kind_column: record.kind,
    }


def _sorted(records: list[Record]) -> list[Record]:
    """日付・種別・キーの順に並べる（毎回同じ並びで出力するため）。"""
    return sorted(records, key=lambda r: (r.date, r.plan_prefix, r.customer_id))
