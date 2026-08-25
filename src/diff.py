"""
src/diff.py — 日付順に並んだファイルを、隣り合う日どうしで突き合わせる

「その日に何件延期になり、何件積み上がったか」を1日ぶんずつ出す。
ある日の値は、その日のファイルと**その直前のファイル**を比べたもの。
土日祝でファイルが飛んでいれば、飛んだ手前のファイルが比較相手になる。

戻り値は ``Counts`` —— 「(対象月, 種別, 判定) をキーに、日付ごとの件数 dict を持つ」
形に直接してある。``DailyDiff`` のような中間データクラスは持たない。
"""

import datetime
import logging
from pathlib import Path
from typing import NamedTuple

from src.source import ColumnRule, Record, read_records

logger = logging.getLogger(__name__)


class Counts(NamedTuple):
    """全ファイル分の集計結果。

    - ``compared_dates``: ファイルが存在した日の集合（空セルとの区別用）
    - ``by_row``: 行キー (対象月, 種別, 判定) → {日付: 件数}。
      該当しない日はキーに含まれない（CSV 側で空セルとして扱う）。
    """

    compared_dates: set[datetime.date]
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]]


STATUS_ADDED = "積み上げ"
STATUS_POSTPONED = "延期"


def compute_counts(
    dated_files: list[tuple[datetime.date, Path]],
    target_months: list[tuple[int, int]],
    plan_prefixes: tuple[str, ...],
    kinds: tuple[str, ...],
    rules: tuple[ColumnRule, ...],
    range_start: datetime.date | None = None,
) -> Counts:
    """日付の古い順に並んだファイルを、隣り合う組で突き合わせ、対象月ごとに数える。

    1ファイルは1回だけ読む。読んだ行は対象月で振り分ける（複数月が対象なら
    そのぶん全部数える）。
    ``range_start`` を渡すと、範囲内のファイルに対して ``(n/total)`` の
    進捗をログに出す。範囲外（比較相手として例外的に読む1ファイル）は
    進捗ログの対象外。
    """
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]] = {}
    compared_dates: set[datetime.date] = set()
    previous_by_month: dict[tuple[int, int], dict[str, Record]] | None = None
    previous_path: Path | None = None

    # 範囲内ファイル数（進捗の分母）。範囲外（比較相手）は含めない
    in_range_total = sum(
        1 for date, _ in dated_files if range_start is None or date >= range_start
    )
    in_range_index = 0

    for date, path in dated_files:
        is_in_range = range_start is None or date >= range_start
        progress: tuple[int, int] | None = None
        if is_in_range and range_start is not None:
            in_range_index += 1
            progress = (in_range_index, in_range_total)

        records = read_records(path, plan_prefixes, kinds, rules, progress=progress)
        # 対象月ごとに「その対象月の行だけ」を取り出した辞書を作る
        current_by_month: dict[tuple[int, int], dict[str, Record]] = {
            month: {} for month in target_months
        }
        for record in records.values():
            month = (record.date.year, record.date.month)
            if month in current_by_month:
                current_by_month[month][record.customer_id] = record

        if previous_by_month is not None:
            for month in target_months:
                previous_records = previous_by_month[month]
                current_records = current_by_month[month]
                added_keys = current_records.keys() - previous_records.keys()
                postponed_keys = previous_records.keys() - current_records.keys()
                for key in added_keys:
                    _inc(by_row, month, current_records[key].plan_prefix, STATUS_ADDED, date)
                for key in postponed_keys:
                    _inc(by_row, month, previous_records[key].plan_prefix, STATUS_POSTPONED, date)
            compared_dates.add(date)
            _log_daily_diff(date, target_months, by_row, previous_path, path)

        previous_by_month = current_by_month
        previous_path = path

    logger.info(
        "%d 日ぶんを比較しました", len(compared_dates)
    )
    return Counts(
        compared_dates=compared_dates,
        by_row=by_row,
    )


def _log_daily_diff(
    date: datetime.date,
    target_months: list[tuple[int, int]],
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]],
    previous_path: Path | None,
    current_path: Path,
) -> None:
    """1日ぶんの集計結果を INFO で出す。対象月が複数のときは対象月ごとにも出す。

    ファイル2つの名前は ``(一覧_YYYYMMDD.xlsx → 一覧_YYYYMMDD.xlsx)`` で後ろに添える。
    対象月が1つのときは ``[YYYY-MM]`` を省略し、複数あるときは各行に付ける。
    """
    show_month_suffix = len(target_months) > 1
    previous_name = previous_path.name if previous_path is not None else "?"
    current_name = current_path.name
    file_pair = f"（{previous_name} → {current_name}）"
    for month in target_months:
        added = _sum_for(by_row, month, STATUS_ADDED, date)
        postponed = _sum_for(by_row, month, STATUS_POSTPONED, date)
        suffix = f" [{month[0]:04d}-{month[1]:02d}]" if show_month_suffix else ""
        logger.info(
            "%s%s: 積み上げ %d 件 / 延期 %d 件%s",
            date.isoformat(),
            suffix,
            added,
            postponed,
            file_pair,
        )


def _sum_for(
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]],
    month: tuple[int, int],
    status: str,
    date: datetime.date,
) -> int:
    """``(対象月, 種別, status)`` の各行について ``date`` の値を合計する。"""
    return sum(
        values.get(date, 0)
        for key, values in by_row.items()
        if key[0] == month and key[2] == status
    )


def _inc(
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]],
    month: tuple[int, int],
    plan_prefix: str,
    status: str,
    date: datetime.date,
) -> None:
    """(対象月, 種別, 判定) のセルに 1 を足す。"""
    key = (month, plan_prefix, status)
    per_date = by_row.setdefault(key, {})
    per_date[date] = per_date.get(date, 0) + 1