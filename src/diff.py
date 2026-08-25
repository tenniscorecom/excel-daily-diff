"""
src/diff.py — 日付順に並んだファイルを、隣り合う日どうしで突き合わせる

「その日に何件延期になり、何件積み上がったか」を1日ぶんずつ出す。
``on_day_done`` を渡すと、日1日ぶんの集計が終わるたびに呼ばれる。
長い処理の途中で落ちても、それまでの計算結果を外に出せるようにするための
フック（``run.py`` が CSV 書き出しに使う）。

入力ファイル ``一覧_YYYYMMDD.xlsx`` は**前日終了時点**のデータ。
``一覧_20260825.xlsx`` の中身は 8/24 終了時点の状態を指し、``一覧_20260824.xlsx``
との差分は「8/24 に動いたぶん」、つまり **業務日 = 前側のファイルの日付**。
ここで言う「業務日 D の列」とは「``一覧_(D+1)`` と ``一覧_(D+2)`` を比べた結果」。

戻り値は ``Counts`` —— 「行キー (対象月, 種別, 判定) をキーに、**業務日**ごとの
件数 dict を持つ」形に直接してある。
"""

import datetime
import logging
from collections.abc import Callable
from pathlib import Path
from typing import NamedTuple

from src.source import Progress, Record, read_records

logger = logging.getLogger(__name__)


# 集計表の 1 行を指すキー。「対象月は "YYYY-MM" の文字列」「種別は1語」「判定は1語」
# の3要素なので NamedTuple にして、``key.target_month`` のように名前で触れるようにした。
# 3重のタプルより読みやすく、辞書キーにもなる（NamedTuple は hashable）。
class RowKey(NamedTuple):
    target_month: str  # "YYYY-MM"（業務日の属する月をそのまま文字列化）
    plan: str
    status: str


# ``by_row`` のキー・値の型。``Counts`` と ``on_day_done`` のシグネチャで
# 共有するため NamedTuple の外に置く。
ByRow = dict[RowKey, dict[datetime.date, int]]

# 1日ぶん集計が終わったあとに呼ばれるフック。``by_row`` と ``compared_dates`` は
# ミュータブルで、ここまでの累積状態をそのまま渡す。書き出し側でファイルに
# 落とせば、途中で例外が出ても途中までが残る。``target_dates_by_month`` は
# その日までに確定した「対象月 → 対象月だった業務日」の対応（書き出し時点で
# 未確定の月は省略される）。
DaySavedCallback = Callable[
    [ByRow, set[datetime.date], dict[str, set[datetime.date]] | None],
    None,
]

# 業務日 → 対象月のリストを返す関数。``run.py`` が ``_target_months`` を渡す。
# 1ファイルは1回だけ読み、そのファイルが担当する業務日（自身 = current、直後 =
# previous）の両方で必要になる対象月すべてに振り分ける。
TargetMonthsFor = Callable[[datetime.date], list[str]]


class Counts(NamedTuple):
    """全ファイル分の集計結果。

    - ``compared_dates``: ファイルが存在した日の集合（空セルとの区別用）
    - ``by_row``: 行キー (対象月, 種別, 判定) → {業務日: 件数}。
      該当しない日はキーに含まれない（CSV 側で空セルとして扱う）
    - ``target_dates_by_month``: 対象月 → その月が対象月だった業務日の集合。
      対象月の行で ``"0"`` と空セルを区別するために ``write_csv`` が参照する
    """

    compared_dates: set[datetime.date]
    by_row: ByRow
    target_dates_by_month: dict[str, set[datetime.date]]


STATUS_ADDED = "積み上げ"
STATUS_POSTPONED = "延期"


def compute_counts(
    dated_files: list[tuple[datetime.date, Path]],
    target_months_for: TargetMonthsFor,
    plan_prefixes: tuple[str, ...],
    kinds: tuple[str, ...],
    rules: tuple,
    range_start: datetime.date | None = None,
    on_day_done: DaySavedCallback | None = None,
) -> Counts:
    """日付の古い順に並んだファイルを、隣り合う組で突き合わせ、対象月ごとに数える。

    1ファイルは1回だけ読む。読んだ行は、**そのファイルが担当する業務日**
    （自身のファイル日付 - 1日 = ``current`` 側、直後のファイル日付 - 1日 =
    ``previous`` 側）のそれぞれが必要とする対象月すべてに振り分ける。

    ``target_months_for`` は業務日 → 対象月のリストを返す関数。``run.py``
    が ``_target_months`` を渡す。

    ``range_start`` を渡すと、範囲内のファイルに対して ``(n/total)`` の
    進捗をログに出す。範囲外（比較相手として例外的に読む1ファイル）は
    進捗ログの対象外。

    ``on_day_done`` を渡すと、日1日ぶんの突き合わせが終わるたびに呼ばれる。
    ``by_row`` と ``compared_dates`` はミュータブルで、ここまでの累積状態を
    そのまま渡す（呼ぶ側でファイルへ書き出せば、途中で落ちても途中までが残る）。
    """
    by_row: ByRow = {}
    compared_dates: set[datetime.date] = set()
    target_dates_by_month: dict[str, set[datetime.date]] = {}
    previous_by_month: dict[str, dict[str, Record]] | None = None
    previous_path: Path | None = None

    # 範囲内ファイル数（進捗の分母）。範囲外（比較相手）は含めない
    in_range_total = sum(
        1 for date, _ in dated_files if range_start is None or date >= range_start
    )
    in_range_index = 0

    # 先に「ファイル日付 → 業務日」を求めておき、各ファイルのバケットに必要な
    # 対象月の集合を隣接2日分から計算する。
    business_dates: list[datetime.date] = [
        date - datetime.timedelta(days=1) for date, _ in dated_files
    ]

    for index, (date, path) in enumerate(dated_files):
        is_in_range = range_start is None or date >= range_start
        progress: Progress | None = None
        if is_in_range and range_start is not None:
            in_range_index += 1
            progress = Progress(in_range_index, in_range_total)

        # このファイルが担当する対象月の集合。
        # - 自身 = current 側 → 業務日 = date - 1日 の対象月
        # - 直後 = previous 側 → 業務日 = dated_files[index+1][0] - 1日 の対象月
        # 最終ファイルは previous 側が無いので current 側だけ。
        relevant_months: set[str] = set(target_months_for(business_dates[index]))
        if index + 1 < len(dated_files):
            relevant_months.update(target_months_for(business_dates[index + 1]))

        records = read_records(path, plan_prefixes, kinds, rules, progress=progress)
        # 対象月ごとに「その対象月の行だけ」を取り出した辞書を作る
        current_by_month: dict[str, dict[str, Record]] = {
            month: {} for month in relevant_months
        }
        for record in records.values():
            month = f"{record.date.year:04d}-{record.date.month:02d}"
            if month in current_by_month:
                current_by_month[month][record.customer_id] = record

        if previous_by_month is not None:
            # 業務日 = ファイル日付 - 1日（入力ファイルは前日終了時点のデータ）。
            # ``compared_dates`` も業務日で持つ。
            business_date = business_dates[index]
            target_months_this_day = target_months_for(business_date)
            for month in target_months_this_day:
                previous_records = previous_by_month.get(month, {})
                current_records = current_by_month.get(month, {})
                added_keys = current_records.keys() - previous_records.keys()
                postponed_keys = previous_records.keys() - current_records.keys()
                for key in added_keys:
                    _inc(by_row, month, current_records[key].plan_prefix, STATUS_ADDED, business_date)
                for key in postponed_keys:
                    _inc(by_row, month, previous_records[key].plan_prefix, STATUS_POSTPONED, business_date)
                # この業務日で対象月だったことを記録。CSV 出力側で「対象月でない業務日は
                # 空セル」にするために使う。
                target_dates_by_month.setdefault(month, set()).add(business_date)
            compared_dates.add(business_date)
            # ログは処理の進行に合わせて直書きする（途中でどこまで進んで
            # いたかがログから追えるように）。
            show_month_suffix = len(target_months_this_day) > 1
            previous_name = previous_path.name if previous_path is not None else "?"
            current_name = path.name
            file_pair = f"（{previous_name} → {current_name}）"
            for month in target_months_this_day:
                added = _sum_for(by_row, month, STATUS_ADDED, business_date)
                postponed = _sum_for(by_row, month, STATUS_POSTPONED, business_date)
                suffix = f" [{month}]" if show_month_suffix else ""
                logger.info(
                    "業務日 %s%s: 積み上げ %d 件 / 延期 %d 件%s",
                    business_date.isoformat(),
                    suffix,
                    added,
                    postponed,
                    file_pair,
                )
            if business_date != date:
                # 業務日とファイル名が1日ずれていることを明示（混同防止）。
                logger.debug(
                    "業務日 %s = ファイル %s の前日終了時点（%s → %s）",
                    business_date.isoformat(),
                    date.isoformat(),
                    previous_name,
                    current_name,
                )
            if on_day_done is not None:
                on_day_done(by_row, compared_dates, target_dates_by_month)

        previous_by_month = current_by_month
        previous_path = path

    logger.info("%d 日ぶんを比較しました", len(compared_dates))
    return Counts(
        compared_dates=compared_dates,
        by_row=by_row,
        target_dates_by_month=target_dates_by_month,
    )


def _sum_for(
    by_row: ByRow,
    month: str,
    status: str,
    date: datetime.date,
) -> int:
    """``(対象月, 種別, status)`` の各行について ``date`` の値を合計する。"""
    return sum(
        values.get(date, 0)
        for key, values in by_row.items()
        if key.target_month == month and key.status == status
    )


def _inc(
    by_row: ByRow,
    month: str,
    plan_prefix: str,
    status: str,
    date: datetime.date,
) -> None:
    """(対象月, 種別, 判定) のセルに 1 を足す。"""
    key = RowKey(month, plan_prefix, status)
    per_date = by_row.setdefault(key, {})
    per_date[date] = per_date.get(date, 0) + 1
