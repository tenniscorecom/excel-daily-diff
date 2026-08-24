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
) -> Counts:
    """日付の古い順に並んだファイルを、隣り合う組で突き合わせ、対象月ごとに数える。

    1ファイルは1回だけ読む。読んだ行は対象月で振り分ける（複数月が対象なら
    そのぶん全部数える）。
    """
    by_row: dict[tuple[tuple[int, int], str, str], dict[datetime.date, int]] = {}
    compared_dates: set[datetime.date] = set()
    previous_by_month: dict[tuple[int, int], dict[str, Record]] | None = None

    for date, path in dated_files:
        records = read_records(path, plan_prefixes, kinds, rules)
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
            logger.debug(
                "%s: 積み上げ合計 %d / 延期合計 %d",
                date,
                sum(
                    sum(values.values())
                    for key, values in by_row.items()
                    if key[2] == STATUS_ADDED and date in values
                ),
                sum(
                    sum(values.values())
                    for key, values in by_row.items()
                    if key[2] == STATUS_POSTPONED and date in values
                ),
            )

        previous_by_month = current_by_month

    logger.info(
        "%d 日ぶんを比較しました", len(compared_dates)
    )
    return Counts(
        compared_dates=compared_dates,
        by_row=by_row,
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