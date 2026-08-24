import datetime
from pathlib import Path
from unittest.mock import patch

from src.diff import STATUS_ADDED, STATUS_POSTPONED, compute_counts
from src.source import Record


def _record(customer_id: str, day: int = 1) -> Record:
    return Record(customer_id, datetime.date(2026, 4, day), "標準")


def test_compute_counts_compares_adjacent_files_and_reads_each_once() -> None:
    paths = [Path(f"一覧_2026042{day}.xlsx") for day in range(3)]
    records = [
        {"a": _record("a", 1), "b": _record("b", 2)},
        {"b": _record("b", 3), "c": _record("c", 4)},
        {"c": _record("c", 5), "d": _record("d", 6)},
    ]
    dated_files = [(datetime.date(2026, 4, 20 + index), path) for index, path in enumerate(paths)]
    target_months = [(2026, 4)]

    with patch("src.diff.read_records", side_effect=records) as reader:
        counts = compute_counts(dated_files, target_months, ("標準",), ("完了",), ())

    added = counts.by_row[(target_months[0], "標準", STATUS_ADDED)]
    postponed = counts.by_row[(target_months[0], "標準", STATUS_POSTPONED)]
    assert added[datetime.date(2026, 4, 21)] == 1  # "c"
    assert added[datetime.date(2026, 4, 22)] == 1  # "d"
    assert postponed[datetime.date(2026, 4, 21)] == 1  # "a"
    assert postponed[datetime.date(2026, 4, 22)] == 1  # "b"
    assert counts.compared_dates == {datetime.date(2026, 4, 21), datetime.date(2026, 4, 22)}
    assert [call.args[0] for call in reader.call_args_list] == paths


def test_compute_counts_only_buckets_into_target_months() -> None:
    paths = [Path(f"一覧_20260{day}.xlsx") for day in (501, 502)]
    # 5/1: 4月と5月のレコードが混在、5/2: 4月のレコードが消えた
    records = [
        {
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
        {
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 5, 1), paths[0]),
        (datetime.date(2026, 5, 2), paths[1]),
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(dated_files, [(2026, 4), (2026, 5)], ("標準",), ("完了",), ())

    # 4月ぶんの行: apr が 5/2 に消えた → 5/2 に「延期 1」
    assert counts.by_row[((2026, 4), "標準", STATUS_POSTPONED)][datetime.date(2026, 5, 2)] == 1
    # 5月の行は「5/2 にも "may" があるので変化なし」→ 集計0件
    assert (2026, 5, "標準", STATUS_ADDED) not in counts.by_row
    assert (2026, 5, "標準", STATUS_POSTPONED) not in counts.by_row


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
        counts = compute_counts(dated_files, [(2026, 4)], ("標準", "上位"), ("完了",), ())

    assert counts.by_row[((2026, 4), "上位", STATUS_POSTPONED)][datetime.date(2026, 4, 22)] == 1
    assert (2026, 4, "標準", STATUS_ADDED) not in counts.by_row