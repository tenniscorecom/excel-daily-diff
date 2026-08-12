import datetime
from pathlib import Path

import pytest
from comken import config

from src.exceptions import InvalidDateSettingError, InvalidMonthSettingError
from src.settings import ColumnRule, _load_rules, _to_date, _to_year_month


def test_to_date_accepts_required_format() -> None:
    assert _to_date("START_DATE", "2026-04-21") == datetime.date(2026, 4, 21)


@pytest.mark.parametrize("value", ["2026/04/21", "2026-02-30", ""])
def test_to_date_raises_for_invalid_setting(value: str) -> None:
    with pytest.raises(InvalidDateSettingError):
        _to_date("START_DATE", value)


def test_to_year_month_accepts_required_format() -> None:
    assert _to_year_month("2026-08") == (2026, 8)


@pytest.mark.parametrize("value", ["2026/08", "2026-8", "2026-13", ""])
def test_to_year_month_raises_for_invalid_setting(value: str) -> None:
    with pytest.raises(InvalidMonthSettingError):
        _to_year_month(value)


def test_load_rules_reads_optional_sections_and_comma_separated_words(
    tmp_path: Path, restore_config_singleton
) -> None:
    path = tmp_path / "rules.ini"
    path.write_text(
        """[INCLUDE_CONTAINS]
地域 = 離島, 山間部

[EXCLUDE_CONTAINS]

[INCLUDE_EXACT]
担当 =
""",
        encoding="utf-8",
    )
    config.read(path)

    assert _load_rules() == (
        ColumnRule("地域", ("離島", "山間部"), is_contains=True, is_exclude=False),
    )
