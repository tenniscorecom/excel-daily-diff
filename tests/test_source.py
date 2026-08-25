import datetime
import logging
from pathlib import Path

import pytest
from comken.exceptions import ExcelColumnNotFoundError

from src.source import ColumnRule, read_records


def test_filters_by_plan_kind_and_skips_out_of_target_rows(
    tmp_path: Path, make_book
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ["first", datetime.date(2026, 4, 1), "標準A", "完了"],
            ["last", datetime.datetime(2026, 4, 30, 12), "上位B", "予定"],
            ["plan_contains", "2026年04月22日", "特別な標準", "完了"],
            ["kind_contains", "2026/04/22 00:00:00", "標準A", "完了予定"],
            ["unknown_plan", "2026-04-22", "その他", "完了"],
        ],
    )

    records = read_records(path, ("標準", "上位"), ("完了", "予定"), ())

    assert set(records) == {"first", "last"}
    assert records["first"].plan_prefix == "標準"
    assert records["last"].plan_prefix == "上位"


def test_skips_empty_and_broken_dates(tmp_path: Path, make_book) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ["empty", None, "標準A", "完了"],
            ["broken", "日付ではない", "標準A", "完了"],
            ["valid", "2026-04-22", "標準A", "完了"],
        ],
    )

    assert set(read_records(path, ("標準",), ("完了",), ())) == {"valid"}


def test_normalizes_numeric_customer_id_and_keeps_first_duplicate(
    tmp_path: Path, make_book
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            [1001, "2026-04-21", "標準A", "完了"],
            ["1001", "2026-04-22", "上位B", "予定"],
        ],
    )

    records = read_records(path, ("標準", "上位"), ("完了", "予定"), ())

    assert list(records) == ["1001"]
    assert records["1001"].date == datetime.date(2026, 4, 21)


@pytest.mark.parametrize(
    ("rule", "expected"),
    [
        (ColumnRule("地域", ("離島",), True, False), {"contains", "exact"}),
        (ColumnRule("地域", ("離島",), True, True), {"other"}),
        (ColumnRule("地域", ("離島",), False, False), {"exact"}),
        (ColumnRule("地域", ("離島",), False, True), {"contains", "other"}),
    ],
)
def test_applies_each_rule_operator(
    tmp_path: Path, make_book, rule: ColumnRule, expected: set[str]
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ["contains", "2026-04-22", "標準A", "完了", "離島A地区"],
            ["exact", "2026-04-22", "標準A", "完了", "離島"],
            ["other", "2026-04-22", "標準A", "完了", "市街地"],
        ],
        ["顧客番号", "予定日", "種別", "状態", "地域"],
    )

    assert set(read_records(path, ("標準",), ("完了",), (rule,))) == expected


def test_multiple_words_are_or_and_multiple_rules_are_and(
    tmp_path: Path, make_book
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ["both_a", "2026-04-22", "標準A", "完了", "離島A地区", "東支店"],
            ["both_b", "2026-04-22", "標準A", "完了", "山間部", "東支店"],
            ["wrong_area", "2026-04-22", "標準A", "完了", "市街地", "東支店"],
            ["wrong_branch", "2026-04-22", "標準A", "完了", "離島", "西支店"],
        ],
        ["顧客番号", "予定日", "種別", "状態", "地域", "支店"],
    )
    rules = (
        ColumnRule("地域", ("離島", "山間部"), True, False),
        ColumnRule("支店", ("東支店",), False, False),
    )

    assert set(read_records(path, ("標準",), ("完了",), rules)) == {
        "both_a",
        "both_b",
    }


def test_raises_when_rule_column_is_missing(tmp_path: Path, make_book) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [["id", "2026-04-22", "標準A", "完了"]],
    )

    with pytest.raises(ExcelColumnNotFoundError):
        read_records(
            path,
            ("標準",),
            ("完了",),
            (ColumnRule("地域", ("離島",), True, True),),
        )


def test_matches_columns_with_middle_dot_and_surrounding_spaces(
    tmp_path: Path, make_book
) -> None:
    """中黒を含む列名で絞り込めること、見出しの前後の空白が邪魔をしないこと。"""
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ["stay", "2026-04-22", "標準A", "完了", "市街地"],
            ["drop", "2026-04-22", "標準A", "完了", "離島B・北地区"],
        ],
        ["顧客番号", "予定日", "種別", "状態", " 住所・地域 "],
    )
    rule = ColumnRule("住所・地域", ("離島",), is_contains=True, is_exclude=True)

    assert set(read_records(path, ("標準",), ("完了",), (rule,))) == {"stay"}


def test_sheet_name_log_includes_used_sheet(
    tmp_path: Path, make_book, caplog: pytest.LogCaptureFixture
) -> None:
    """使ったシート名がログに出る。"""
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [["a", "2026-04-22", "標準A", "完了"]],
    )

    with caplog.at_level(logging.INFO):
        read_records(path, ("標準",), ("完了",), ())

    # "[Sheet1]" の形式で使われるシート名がログに出る
    used_sheet_logs = [record.message for record in caplog.records if "[Sheet1]" in record.message]
    assert any("条件に合う行" in message for message in used_sheet_logs)
