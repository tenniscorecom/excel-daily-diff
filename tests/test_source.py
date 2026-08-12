import datetime
from pathlib import Path

import pytest
from comken.exceptions import ExcelColumnNotFoundError

from src.settings import ColumnRule, Criteria, SourceLayout
from src.source import read_records


def _with_rules(criteria: Criteria, *rules: ColumnRule) -> Criteria:
    return Criteria(
        criteria.target_year,
        criteria.target_month,
        criteria.plan_prefixes,
        criteria.kinds,
        rules,
    )


def test_filters_by_all_three_conditions_and_target_month_boundaries(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ("previous_month", datetime.date(2026, 3, 31), "標準A", "完了"),
            ("first", datetime.date(2026, 4, 1), "標準A", "完了"),
            ("last", datetime.datetime(2026, 4, 30, 12), "上位B", "予定"),
            ("next_month", "2026/05/01", "標準A", "完了"),
            ("plan_contains", "2026年04月22日", "特別な標準", "完了"),
            ("kind_contains", "2026/04/22 00:00:00", "標準A", "完了予定"),
        ],
    )

    records = read_records(path, layout, criteria)

    assert set(records) == {"first", "last"}
    assert records["first"].plan_prefix == "標準"
    assert records["last"].plan_prefix == "上位"


def test_skips_empty_and_broken_dates(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ("empty", None, "標準A", "完了"),
            ("broken", "日付ではない", "標準A", "完了"),
            ("valid", "2026-04-22", "標準A", "完了"),
        ],
    )

    assert set(read_records(path, layout, criteria)) == {"valid"}


def test_normalizes_numeric_customer_id_and_keeps_first_duplicate(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            (1001, "2026-04-21", "標準A", "完了"),
            ("1001", "2026-04-22", "上位B", "予定"),
        ],
    )

    records = read_records(path, layout, criteria)

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
    tmp_path: Path,
    make_book,
    layout: SourceLayout,
    criteria: Criteria,
    rule: ColumnRule,
    expected: set[str],
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ("contains", "2026-04-22", "標準A", "完了", "離島A地区"),
            ("exact", "2026-04-22", "標準A", "完了", "離島"),
            ("other", "2026-04-22", "標準A", "完了", "市街地"),
        ],
        ["顧客番号", "予定日", "種別", "状態", "地域"],
    )

    assert set(read_records(path, layout, _with_rules(criteria, rule))) == expected


def test_multiple_words_are_or_and_multiple_rules_are_and(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ("both_a", "2026-04-22", "標準A", "完了", "離島A地区", "東支店"),
            ("both_b", "2026-04-22", "標準A", "完了", "山間部", "東支店"),
            ("wrong_area", "2026-04-22", "標準A", "完了", "市街地", "東支店"),
            ("wrong_branch", "2026-04-22", "標準A", "完了", "離島", "西支店"),
        ],
        ["顧客番号", "予定日", "種別", "状態", "地域", "支店"],
    )
    rules = (
        ColumnRule("地域", ("離島", "山間部"), True, False),
        ColumnRule("支店", ("東支店",), False, False),
    )

    assert set(read_records(path, layout, _with_rules(criteria, *rules))) == {
        "both_a",
        "both_b",
    }


def test_raises_when_rule_column_is_missing(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [("id", "2026-04-22", "標準A", "完了")],
    )
    criteria = _with_rules(criteria, ColumnRule("地域", ("離島",), True, True))

    with pytest.raises(ExcelColumnNotFoundError):
        read_records(path, layout, criteria)


def test_matches_columns_with_middle_dot_and_surrounding_spaces(
    tmp_path: Path, make_book, layout: SourceLayout, criteria: Criteria
) -> None:
    """中黒を含む列名で絞り込めること、見出しの前後の空白が邪魔をしないこと。"""
    path = make_book(
        tmp_path / "一覧_20260423.xlsx",
        [
            ("stay", "2026-04-22", "標準A", "完了", "市街地"),
            ("drop", "2026-04-22", "標準A", "完了", "離島B・北地区"),
        ],
        ["顧客番号", "予定日", "種別", "状態", " 住所・地域 "],
    )
    rule = ColumnRule("住所・地域", ("離島",), is_contains=True, is_exclude=True)

    assert set(read_records(path, layout, _with_rules(criteria, rule))) == {"stay"}
