import datetime
from pathlib import Path
from unittest.mock import patch

import pytest
from comken import Config
from openpyxl import Workbook

from src.source import ColumnRule, read_records

HEADERS = ["顧客番号", "予定日", "種別", "状態"]


@pytest.fixture
def config_for_tests(monkeypatch, tmp_path):
    """テスト用の Config(path) を作って、``comken.config`` の束縛先を上書きする。"""

    def _use(content: str) -> Config:
        path = tmp_path / "config.ini"
        path.write_text(content, encoding="utf-8")
        test_config = Config(path)
        # src.source / src.run は ``from comken import config`` で束縛した名前を
        # 持っているため、``comken.config`` を差し替えてもそちらには届かない。
        # それぞれのモジュール側の名前も差し替える。
        monkeypatch.setattr("src.source.config", test_config)
        monkeypatch.setattr("src.run.config", test_config)
        return test_config

    return _use


@pytest.fixture
def plan_prefixes() -> tuple[str, ...]:
    return ("標準", "上位")


@pytest.fixture
def kinds() -> tuple[str, ...]:
    return ("完了", "予定")


@pytest.fixture
def make_book():
    def _make_book(
        path: Path,
        rows: list[list[object]],
        headers: list[object] | None = None,
    ) -> Path:
        workbook = Workbook()
        worksheet = workbook.active
        assert worksheet is not None
        worksheet.title = "Sheet1"
        worksheet.append(headers if headers is not None else HEADERS)
        for row in rows:
            worksheet.append(list(row))
        workbook.save(path)
        workbook.close()
        return path

    return _make_book


_CONFIG_FOR_READ_RECORDS = """[FILES]
INPUT_FOLDER = C:\\作業\\input
FILE_PATTERN = 一覧_*.xlsx

[SOURCE]
SHEET_NAME = Sheet1
HEADER_ROW = 1
KEY_COLUMN = 顧客番号
DATE_COLUMN = 予定日
PLAN_COLUMN = 種別
KIND_COLUMN = 状態

[FILTER]
PLAN_PREFIXES = [標準, 上位]
KINDS = [完了, 予定]
"""


@pytest.fixture(autouse=True)
def _patch_config_for_read_records(config_for_tests) -> None:
    """``read_records`` は config を直接読むため、テストごとに設定を入れておく。"""
    config_for_tests(_CONFIG_FOR_READ_RECORDS)