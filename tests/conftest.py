from collections.abc import Iterable
from pathlib import Path

import pytest
from comken import config
from openpyxl import Workbook

from src.settings import Criteria, SourceLayout

HEADERS = ["顧客番号", "予定日", "種別", "状態"]


@pytest.fixture
def layout() -> SourceLayout:
    return SourceLayout("Sheet1", 1, "顧客番号", "予定日", "種別", "状態")


@pytest.fixture
def criteria() -> Criteria:
    return Criteria(2026, 4, ("標準", "上位"), ("完了", "予定"))


@pytest.fixture
def make_book():
    def _make_book(
        path: Path,
        rows: Iterable[Iterable[object]],
        headers: Iterable[object] = HEADERS,
    ) -> Path:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Sheet1"
        sheet.append(list(headers))
        for row in rows:
            sheet.append(list(row))
        workbook.save(path)
        workbook.close()
        return path

    return _make_book


@pytest.fixture
def restore_config_singleton():
    """一時設定を読むテストの後で comken.config の共有状態を元に戻す。"""
    original = config._singleton
    try:
        yield
    finally:
        config._singleton = original
