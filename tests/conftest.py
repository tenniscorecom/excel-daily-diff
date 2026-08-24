from collections.abc import Iterable
from pathlib import Path

import pytest
from comken import Config
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
def use_config(monkeypatch):
    """テスト用の Config(path) を ``src.settings`` から見える位置に差し込む。

    旧 comken は ``config._singleton`` を更新する ``config.read(path)`` だったが、
    新 comken は Config() を毎回生成する遅延シングルトンなので、テストでは
    ``src.settings.config`` 自体を新しいインスタンスへ差し替える。
    """

    def _use(path: Path) -> Config:
        test_config = Config(path)
        # ``src.settings`` の ``from comken import config`` の束縛先を上書き
        monkeypatch.setattr("src.settings.config", test_config)
        return test_config

    return _use
