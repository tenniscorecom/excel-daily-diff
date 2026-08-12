import datetime
from pathlib import Path
from unittest.mock import patch

from src.diff import daily_diffs
from src.settings import Criteria, SourceLayout
from src.source import Record


def _record(customer_id: str) -> Record:
    return Record(customer_id, datetime.date(2026, 4, 1), "標準", "標準A", "完了")


def test_daily_diffs_compares_adjacent_files_and_reads_each_once(
    layout: SourceLayout, criteria: Criteria
) -> None:
    paths = [Path(f"一覧_2026042{day}.xlsx") for day in range(3)]
    records = [
        {"a": _record("a"), "b": _record("b")},
        {"b": _record("b"), "c": _record("c")},
        {"c": _record("c"), "d": _record("d")},
    ]
    dated_files = [(datetime.date(2026, 4, 20 + index), path) for index, path in enumerate(paths)]

    with patch("src.diff.read_records", side_effect=records) as reader:
        diffs = daily_diffs(dated_files, layout, criteria)

    assert [[record.customer_id for record in diff.added] for diff in diffs] == [["c"], ["d"]]
    assert [[record.customer_id for record in diff.postponed] for diff in diffs] == [["a"], ["b"]]
    assert [call.args[0] for call in reader.call_args_list] == paths
