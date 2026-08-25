import datetime
from pathlib import Path
from unittest.mock import patch

from src.diff import STATUS_ADDED, STATUS_POSTPONED, compute_counts
from src.source import Record


def _record(customer_id: str, day: int = 1) -> Record:
    return Record(customer_id, datetime.date(2026, 4, day), "標準")


def _target_months_for(business_date: datetime.date) -> list[tuple[int, int]]:
    """テスト用の対象月決定。業務日が 23日以降なら翌月も含む（実コードと同じ規約）。"""
    months = [(business_date.year, business_date.month)]
    if business_date.day >= 23:
        if business_date.month == 12:
            months.append((business_date.year + 1, 1))
        else:
            months.append((business_date.year, business_date.month + 1))
    return months


def test_compute_counts_compares_adjacent_files_and_reads_each_once() -> None:
    paths = [Path(f"一覧_2026042{day}.xlsx") for day in range(3)]
    records = [
        {"a": _record("a", 1), "b": _record("b", 2)},
        {"b": _record("b", 3), "c": _record("c", 4)},
        {"c": _record("c", 5), "d": _record("d", 6)},
    ]
    dated_files = [(datetime.date(2026, 4, 20 + index), path) for index, path in enumerate(paths)]

    with patch("src.diff.read_records", side_effect=records) as reader:
        counts = compute_counts(dated_files, _target_months_for, ("標準",), ("完了",), ())

    added = counts.by_row[((2026, 4), "標準", STATUS_ADDED)]
    postponed = counts.by_row[((2026, 4), "標準", STATUS_POSTPONED)]
    # 業務日 = ファイル日付 - 1日。
    # 一覧_20260421.xlsx(業務日 4/20) vs 一覧_20260420.xlsx(業務日 4/19):
    #   c は 4/20 にのみ存在 → 業務日 4/20 に積み上げ 1
    #   a は 4/19 にのみ存在 → 業務日 4/20 に延期 1
    assert added[datetime.date(2026, 4, 20)] == 1  # "c"
    assert postponed[datetime.date(2026, 4, 20)] == 1  # "a"
    # 一覧_20260422.xlsx(業務日 4/21) vs 一覧_20260421.xlsx(業務日 4/20):
    #   d は 4/21 にのみ存在 → 業務日 4/21 に積み上げ 1
    #   b は 4/20 にのみ存在 → 業務日 4/21 に延期 1
    assert added[datetime.date(2026, 4, 21)] == 1  # "d"
    assert postponed[datetime.date(2026, 4, 21)] == 1  # "b"
    assert counts.compared_dates == {datetime.date(2026, 4, 20), datetime.date(2026, 4, 21)}
    assert [call.args[0] for call in reader.call_args_list] == paths


def test_compute_counts_only_buckets_into_target_months() -> None:
    paths = [Path(f"一覧_20260{day}.xlsx") for day in (430, 501, 502)]
    # 4/30: 4月のレコードのみ、5/1: 4月と5月のレコードが混在、5/2: 5月のレコードのみ
    records = [
        {
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
        },
        {
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
        {
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),
        (datetime.date(2026, 5, 1), paths[1]),
        (datetime.date(2026, 5, 2), paths[2]),
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("標準",), ("完了",), ()
        )

    # 業務日 4/30 (5/1 ファイル vs 4/30 ファイル):
    # 4/30 は day 30 >= 23 なので対象月は 4月 と 5月
    #   4月: apr が両方にいる → 差分なし
    #   5月: may が current にのみ存在 → 5月 積み上げ 1
    assert counts.by_row[((2026, 5), "標準", STATUS_ADDED)][datetime.date(2026, 4, 30)] == 1
    # 業務日 5/1 (5/2 ファイル vs 5/1 ファイル):
    # 5/1 は day 1 < 23 なので対象月は 5月のみ
    #   5月: may が両方にいる → 差分なし
    assert ((2026, 4), "標準", STATUS_ADDED) not in counts.by_row
    assert ((2026, 4), "標準", STATUS_POSTPONED) not in counts.by_row
    assert ((2026, 5), "標準", STATUS_ADDED) not in counts.by_row.get(
        ((2026, 5), "標準", STATUS_ADDED), {}
    ) or datetime.date(2026, 5, 1) not in counts.by_row[((2026, 5), "標準", STATUS_ADDED)]


def test_compute_counts_groups_by_plan_prefix() -> None:
    paths = [Path("一覧_20260421.xlsx"), Path("一覧_20260422.xlsx")]
    records = [
        {
            "a": Record("a", datetime.date(2026, 4, 10), "標準"),
            "b": Record("b", datetime.date(2026, 4, 10), "上位"),
        },
        {
            "a": Record("a", datetime.date(2026, 4, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 21), paths[0]),
        (datetime.date(2026, 4, 22), paths[1]),
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("標準", "上位"), ("完了",), ()
        )

    # 一覧_20260421.xlsx(業務日 4/20) vs 一覧_20260422.xlsx(業務日 4/21):
    # b が消えた → 業務日 4/21 に 上位 の延期 1
    assert counts.by_row[((2026, 4), "上位", STATUS_POSTPONED)][datetime.date(2026, 4, 21)] == 1
    assert ((2026, 4), "標準", STATUS_ADDED) not in counts.by_row


def test_compute_counts_uses_business_date_target_months_across_month_boundary() -> None:
    """月をまたぐ業務日（4/30 → 5/1）で、両方の業務日の対象月が和集合でバケットされる。

    業務日 4/30 → 4月対象、業務日 5/1 → 5月対象。一覧_20260501.xlsx は両方の
    突き合わせで previous/current になるので、4月と5月の両方でバケットされる。
    """
    paths = [Path("一覧_20260430.xlsx"), Path("一覧_20260501.xlsx")]
    records = [
        {
            # 4月のレコード
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
        },
        {
            # 5月のレコード
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),  # → 業務日 4/29
        (datetime.date(2026, 5, 1), paths[1]),   # → 業務日 4/30（前）と 5/1（後）
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("標準",), ("完了",), ()
        )

    # 業務日 4/29 (4/30 ファイル vs previous=4/30以前のprevious無し)
    # → 1回目なので previous_by_month が None、比較は走らない
    # 業務日 4/30 (5/1 ファイル vs 4/30 ファイル) → previous=apr、current=may
    # 4月対象では apr が current にない → 4月 延期 1
    assert counts.by_row[((2026, 4), "標準", STATUS_POSTPONED)][datetime.date(2026, 4, 30)] == 1
    # 5月対象では may が current にある、previous にも 4/30ファイルを5月でもバケット
    # していれば両方にあるが、4/30 ファイルには 5月のレコードが無い
    # → current の 5月に may のみ → 5月 積み上げ 1
    assert counts.by_row[((2026, 5), "標準", STATUS_ADDED)][datetime.date(2026, 4, 30)] == 1


def test_compute_counts_reads_each_file_only_once_across_month_boundary() -> None:
    """月をまたぐ日の同じファイルが、previous/current の両方で1回しか読まれない。"""
    paths = [
        Path("一覧_20260430.xlsx"),
        Path("一覧_20260501.xlsx"),
        Path("一覧_20260502.xlsx"),
    ]
    records = [
        {"a": Record("a", datetime.date(2026, 4, 30), "標準")},
        {"a": Record("a", datetime.date(2026, 4, 30), "標準"), "b": Record("b", datetime.date(2026, 5, 1), "標準")},
        {"a": Record("a", datetime.date(2026, 4, 30), "標準"), "b": Record("b", datetime.date(2026, 5, 1), "標準")},
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),  # → 業務日 4/29, 4/30
        (datetime.date(2026, 5, 1), paths[1]),   # → 業務日 4/30, 5/1
        (datetime.date(2026, 5, 2), paths[2]),   # → 業務日 5/1
    ]

    with patch("src.diff.read_records", side_effect=records) as reader:
        compute_counts(dated_files, _target_months_for, ("標準",), ("完了",), ())

    # 各ファイルが読まれるのは1回だけ
    opened = [call.args[0] for call in reader.call_args_list]
    assert opened == paths
    assert len(opened) == len(set(opened))
