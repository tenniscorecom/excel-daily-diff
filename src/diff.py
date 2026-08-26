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

戻り値は ``Counts`` —— 「行キー (当月/来月ラベル, 種別, 判定) をキーに、
**業務日**ごとの件数 dict を持つ」形に直接してある。
"""

import datetime
import logging
from collections.abc import Callable
from pathlib import Path
from typing import NamedTuple

from src.source import Progress, Record, read_records

logger = logging.getLogger(__name__)


# 出力CSV の行ラベル。対象月の実際の日付文字列 ("2026-08" 等) ではなく、
# 業務日の属する月を「当月」と、翌月を「来月」と相対ラベルで表す。
# 内部のレコード絞り込み（``current_by_month`` / ``previous_by_month`` のキー）には
# 実際の暦月文字列を引き続き使う（案件側の日付と照合する必要があるため）。
LABEL_CURRENT_MONTH = "当月"
LABEL_NEXT_MONTH = "来月"


# 集計表の 1 行を指すキー。「業務日基準の対象月（**当月/来月のラベル**）」、
# 「種別（標準/上位）」、「判定」の3要素NamedTuple。
# ``key.target_month`` のように名前で触れるようにした。
# 2重のタプルより読みやすく、辞書キーにもなる（NamedTuple は hashable）。
#
# フィールド名 ``target_month`` は historic に「対象月」を指すが、中身は
# ``LABEL_CURRENT_MONTH`` / ``LABEL_NEXT_MONTH`` のいずれかになる
# （実際の暦月 "YYYY-MM" ではない）。
class RowKey(NamedTuple):
    target_month: str  # "当月" / "来月"（実際の暦月ではない）
    plan: str  # 種別（標準/上位等）
    status: str


# ``by_row`` のキー・値の型。``Counts`` と ``on_day_done`` のシグネチャで
# 共有するため NamedTuple の外に置く。
# キーは ``(対象月ラベル, 種別, 判定)`` の素のタプルにする（``RowKey`` NamedTuple だと
# キーアクセス時の型共変性が崩れて呼び出し側で Literal 警告が出るため）。
# 値は名称の都合上 NamedTuple ``RowKey`` も用意しておくが、内部のキー操作は
# タプルで行う。
ByRow = dict[tuple[str, str, str], dict[datetime.date, int]]

# 1日ぶん集計が終わったあとに呼ばれるフック。``by_row`` と ``compared_dates`` は
# ミュータブルで、ここまでの累積状態をそのまま渡す。書き出し側でファイルに
# 落とせば、途中で例外が出ても途中までが残る。``target_dates_by_month`` は
# その日までに確定した「対象月のラベル → その対象月だった業務日」の対応
# （書き出し時点で未確定のラベルは省略される）。
DaySavedCallback = Callable[
    [ByRow, set[datetime.date], dict[str, set[datetime.date]] | None],
    None,
]

# 業務日 → 対象月のリスト（実際の暦月文字列）を返す関数。``run.py`` が
# ``_target_months`` を渡す。1ファイルは1回だけ読み、そのファイルが担当する
# 業務日（自身 = current、直後 = previous）の両方で必要になる対象月すべてに
# 振り分ける。
TargetMonthsFor = Callable[[datetime.date], list[str]]


class Counts(NamedTuple):
    """全ファイル分の集計結果。

    - ``compared_dates``: ファイルが存在した日の集合（空セルとの区別用）
    - ``by_row``: 行キー (当月/来月ラベル, 種別, 判定) の ``tuple`` → {業務日: 件数}。
      該当しない日はキーに含まれない（CSV 側で空セルとして扱う）
    - ``target_dates_by_month``: ラベル → そのラベルが対象だった業務日の集合。
      ラベル基準になったので、``write_csv`` が業務日と当月/来月の対応付けに使う
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

    ``target_months_for`` は業務日 → 対象月（実際の暦月文字列）のリストを返す
    関数。``run.py`` が ``_target_months`` を渡す。受け取った暦月文字列を
    ``_target_months`` が返すと保証している順序（当月=0番目、来月=1番目）
    にしたがって、出力用の RowKey / ``target_dates_by_month`` のキーには
    ``当月`` / ``来月`` のラベルを入れて組み立てる。レコード絞り込み用の
    ``current_by_month`` / ``previous_by_month`` のキーには引き続き実際の
    暦月文字列をそのまま使う（案件の日付と照合する必要があるため）。

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
        # 対象月ごとに「その対象月の行だけ」を取り出した辞書を作る。
        # キーは実際の暦月文字列（案件の日付と照合するため）。
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
            # 当該業務日の対象月（実際の暦月文字列）のリスト。
            # ``_target_months`` の出力順序は「業務日自身の月（0番目）= 当月」、
            # 「翌月（1番目）= 来月」が保証されている（``run.py`` 側の実装参照）。
            target_months_this_day = target_months_for(business_date)
            for month_index, month in enumerate(target_months_this_day):
                label = _label_for_month_index(month_index)
                previous_records = previous_by_month.get(month, {})
                current_records = current_by_month.get(month, {})
                added_keys = current_records.keys() - previous_records.keys()
                postponed_keys = previous_records.keys() - current_records.keys()
                for key in added_keys:
                    _inc(by_row, label, current_records[key].plan_prefix,
                         STATUS_ADDED, business_date)
                for key in postponed_keys:
                    _inc(by_row, label, previous_records[key].plan_prefix,
                         STATUS_POSTPONED, business_date)
                # この業務日でそのラベルが対象だったことを記録。CSV 出力側で
                # 「対象ラベルでない業務日は空セル」にするために使う。
                target_dates_by_month.setdefault(label, set()).add(business_date)
            compared_dates.add(business_date)
            # ログは処理の進行に合わせて直書きする（途中でどこまで進んで
            # いたかがログから追えるように）。実際の暦月文字列を見せるのは
            # 運用者がログから実月を追えるようにするため（ラベルのままだと
            # どの月の集計か分からなくなる）。
            show_month_suffix = len(target_months_this_day) > 1
            previous_name = previous_path.name if previous_path is not None else "?"
            current_name = path.name
            file_pair = f"（{previous_name} → {current_name}）"
            for month_index, month in enumerate(target_months_this_day):
                label = _label_for_month_index(month_index)
                added = _sum_for(by_row, label, STATUS_ADDED, business_date)
                postponed = _sum_for(by_row, label, STATUS_POSTPONED, business_date)
                suffix = f" [{month}]" if show_month_suffix else ""
                logger.info(
                    "業務日 %s（%s）%s: 積み上げ %d 件 / 延期 %d 件%s",
                    business_date.isoformat(),
                    label,
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


def _label_for_month_index(index: int) -> str:
    """``_target_months`` の戻り値リスト内の位置を相対ラベルに変換する。

    0 番目 = 業務日自身の月 = ``当月``、1 番目 = 翌月 = ``来月``。
    2 番目以降の実装には依存しない（``_target_months`` の現状は最大 1 個目まで）。
    """
    if index == 0:
        return LABEL_CURRENT_MONTH
    if index == 1:
        return LABEL_NEXT_MONTH
    raise IndexError(
        f"_target_months が {index + 1} 個目の月を返しました"
        "（対応するラベルが定義されていません）"
    )


def _sum_for(
    by_row: ByRow,
    label: str,
    status: str,
    date: datetime.date,
) -> int:
    """``(対象月ラベル, 種別, status)`` の各行のうち ``date`` の値を
    種別をまたいで合計する。

    ログ表示用（「積み上げ合計◯件」相当）で、種別を区別せずに合算した件数を
    出す。キー比較は ``(対象月ラベル, *, 判定)`` の形に揃える（``*`` は種別
    によらず一致する）。
    """
    return sum(
        values.get(date, 0)
        for key, values in by_row.items()
        if key[0] == label and key[2] == status
    )


def _inc(
    by_row: ByRow,
    label: str,
    plan_prefix: str,
    status: str,
    date: datetime.date,
) -> None:
    """(対象月ラベル, 種別, 判定) のセルに 1 を足す。

    種別（標準/上位）ごとに別セルで数える。``plan_prefixes`` の設定順に並ぶ。
    """
    key = (label, plan_prefix, status)
    per_date = by_row.setdefault(key, {})
    per_date[date] = per_date.get(date, 0) + 1
