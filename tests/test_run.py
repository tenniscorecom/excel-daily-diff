import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.exceptions import ComparisonFileNotEnoughError
from src.run import run
from src.settings import Criteria, Settings, SourceLayout


def _settings(
    folder: Path, layout: SourceLayout, criteria: Criteria, output_folder: Path | None = None
) -> Settings:
    return Settings(folder, "一覧_*.xlsx", output_folder or folder, layout, criteria)


def test_run_detects_postponed_and_added_and_writes_report(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    make_book(
        tmp_path / "一覧_20260421.xlsx",
        [
            (1001, "2026-04-21", "標準A", "完了"),
            ("postponed", "2026-04-22", "上位B", "予定"),
        ],
    )
    make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ("1001", "2026-04-21", "標準A", "完了"),
            ("added", "2026-04-23", "標準C", "予定"),
        ],
    )

    output_path = run(_settings(tmp_path, layout, criteria))

    assert output_path.name == "延期積上集計_20260421_20260423.xlsx"
    workbook = load_workbook(output_path, data_only=True)
    summary = workbook["集計"]
    assert summary.max_column == 1 + 3 * 2
    assert [summary.cell(3, column).value for column in range(2, 8)] == [0, 0, 1, 0, 0, 0]
    assert [summary.cell(4, column).value for column in range(2, 8)] == [0, 0, 0, 0, 1, 0]
    detail = list(workbook["明細"].values)
    assert detail == [
        ("判定", "顧客番号", "予定日", "種別", "状態"),
        ("積み上げ", "added", datetime.datetime(2026, 4, 23), "標準C", "予定"),
        ("延期", "postponed", datetime.datetime(2026, 4, 22), "上位B", "予定"),
    ]
    workbook.close()


def test_run_raises_when_only_one_comparison_file_exists(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    make_book(tmp_path / "一覧_20260423.xlsx", [])

    with pytest.raises(ComparisonFileNotEnoughError):
        run(_settings(tmp_path, layout, criteria))
