import datetime
from pathlib import Path
from unittest.mock import patch

import pytest
from comken.toolbox.csv import CSV

from src.run import OUTPUT_NAME, _range_floor, _target_months, run


def _config_text(input_folder: Path, output_folder: Path) -> str:
    """テスト用の config.ini 本文。"""
    return f"""[FILES]
INPUT_FOLDER = {input_folder}
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

[REPORT]
OUTPUT_FOLDER = {output_folder}
"""


@pytest.fixture
def setup_run(config_for_tests):
    """input/output フォルダを tmp に作り、config.ini を読み込む。"""

    def _setup(
        tmp_path: Path, *, kinds_value: str = "[完了, 予定]"
    ) -> tuple[Path, Path]:
        input_folder = tmp_path / "input"
        output_folder = tmp_path / "output"
        input_folder.mkdir()
        output_folder.mkdir()
        config_for_tests(_config_text(input_folder, output_folder).replace(
            "KINDS = [完了, 予定]", f"KINDS = {kinds_value}"
        ))
        return input_folder, output_folder

    return _setup


def test_target_months_uses_only_current_month_before_the_23rd() -> None:
    assert _target_months(datetime.date(2026, 4, 10)) == [(2026, 4)]


def test_target_months_adds_next_month_from_the_23rd_onwards() -> None:
    assert _target_months(datetime.date(2026, 4, 23)) == [(2026, 4), (2026, 5)]


def test_target_months_wraps_year_boundary_in_december() -> None:
    assert _target_months(datetime.date(2026, 12, 25)) == [(2026, 12), (2027, 1)]


def test_run_starts_from_oldest_file_when_no_csv_exists(
    tmp_path: Path, make_book, setup_run
) -> None:
    input_folder, output_folder = setup_run(tmp_path)
    # 4/20 の比較相手として 2/28 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20260228.xlsx", [["stay", "2026-04-10", "標準A", "完了"]])
    make_book(
        input_folder / "一覧_20260420.xlsx",
        [["stay", "2026-04-10", "標準A", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260422.xlsx",
        [["stay", "2026-04-10", "標準A", "完了"], ["add", "2026-04-12", "標準C", "予定"]],
    )
    make_book(
        input_folder / "一覧_20260423.xlsx",
        [["stay", "2026-04-10", "標準A", "完了"]],
    )

    run(datetime.date(2026, 4, 25))

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    target_months = {row["対象月"] for row in rows}
    assert "2026-04" in target_months
    # 4/22: 積み上げ 1, 延期 0
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-22") == "1"
    assert _cell(rows, "2026-04", "標準", "延期", "2026-04-22") == "0"
    # 4/23: 積み上げ 0, 延期 1
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-23") == "0"
    assert _cell(rows, "2026-04", "標準", "延期", "2026-04-23") == "1"


def test_run_writes_blank_for_days_without_files(
    tmp_path: Path, make_book, setup_run
) -> None:
    input_folder, output_folder = setup_run(tmp_path)
    # 4/20 の比較相手として 2/28 を置く
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    # 4/22 のファイルが無い
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    run(datetime.date(2026, 4, 25))

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    # 4/22 列は 0 ではなく空
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-22") == ""
    # 4/21 と 4/23 はそれぞれ 0（変化なし）
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-21") == "0"
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-23") == "0"


def test_run_continues_from_last_csv_date_and_keeps_existing_columns(
    tmp_path: Path, make_book, setup_run
) -> None:
    input_folder, output_folder = setup_run(tmp_path)
    # 1回目: 2/28（比較相手）, 4/20, 4/21 を実行
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run(datetime.date(2026, 4, 25))
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    first_4_21 = _cell(first_rows, "2026-04", "標準", "積み上げ", "2026-04-21")

    # 2回目: 4/22, 4/23 のファイルを追加して再実行
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run(datetime.date(2026, 4, 25))

    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())
    # 4/21 の値が変わらない
    assert _cell(second_rows, "2026-04", "標準", "積み上げ", "2026-04-21") == first_4_21
    # 4/22, 4/23 が追加される
    assert "2026-04-22" in second_rows[0]
    assert "2026-04-23" in second_rows[0]
    # 同じキー (2026-04, 標準, 積み上げ) の行が1つだけ（マージされている）
    matching = [r for r in second_rows
                if r["対象月"] == "2026-04" and r["種別"] == "標準" and r["判定"] == "積み上げ"]
    assert len(matching) == 1


def test_run_rebuilds_full_period_when_conditions_change(
    tmp_path: Path, make_book, config_for_tests
) -> None:
    input_folder = tmp_path / "input"
    output_folder = tmp_path / "output"
    input_folder.mkdir()
    output_folder.mkdir()
    config_for_tests(_config_text(input_folder, output_folder))

    # 1回目: 全ファイルを実行（KINDS = [完了, 予定]）。2/28 は比較相手
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run(datetime.date(2026, 4, 25))
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    first_4_21 = _cell(first_rows, "2026-04", "標準", "積み上げ", "2026-04-21")

    # config.ini を書き換える（KINDS を変える）→ 全期間を作り直す
    config_for_tests(_config_text(input_folder, output_folder).replace(
        "KINDS = [完了, 予定]", "KINDS = [完了]"
    ))
    run(datetime.date(2026, 4, 25))

    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())
    # 4/21 の値がそのまま残る（再計算された結果）
    assert _cell(second_rows, "2026-04", "標準", "積み上げ", "2026-04-21") == first_4_21


def test_run_keeps_old_target_month_rows_with_empty_new_columns(
    tmp_path: Path, make_book, setup_run
) -> None:
    """先月が対象月の行は、新しい日付列が空のまま残る。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 4月のとき 4月のみが対象。比較相手用に 2/28 を置く
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run(datetime.date(2026, 4, 25))
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    # 4月の行が生成されている
    assert any(row["対象月"] == "2026-04" for row in first_rows)

    # 5月以前の対象月のファイルを残しつつ 23日以降の日付で再実行
    # → 5月が対象月に加わる。4月の集計対象はそのまま。
    # このテストでは 4月のレコードが消える（＝延期になる）形で
    # 「古い対象月」4月の行の列も更新されることを確認する。
    make_book(
        input_folder / "一覧_20260422.xlsx",
        [["b", "2026-04-11", "標準B", "予定"]],
    )
    make_book(
        input_folder / "一覧_20260423.xlsx",
        [["b", "2026-04-11", "標準B", "予定"]],
    )
    run(datetime.date(2026, 5, 25))  # 5月25日 → 5月と6月が対象。4/20〜4/23 が下限(4/1)より後

    # 4月の行が残っていること、5月と6月の行が新たに作られていることを確認
    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())

    # 4月の行が残る
    assert any(row["対象月"] == "2026-04" for row in second_rows)
    # 5月と6月の行が新たに作られている
    assert any(row["対象月"] == "2026-05" for row in second_rows)
    assert any(row["対象月"] == "2026-06" for row in second_rows)
    # 4月の行の 4/22, 4/23 列は空（4月はもう対象月でないため）
    row_apr_added = _row(second_rows, "2026-04", "標準", "積み上げ")
    assert row_apr_added["2026-04-22"] == ""
    assert row_apr_added["2026-04-23"] == ""


def test_run_targets_two_months_from_the_23rd(
    tmp_path: Path, make_book, setup_run
) -> None:
    input_folder, output_folder = setup_run(tmp_path)
    # 範囲内最初（4/10）の比較相手として 2/28 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260410.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260418.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    # 5月と4月のレコードが混在するファイル
    make_book(
        input_folder / "一覧_20260420.xlsx",
        [
            ["apr", "2026-04-10", "標準A", "完了"],
            ["may", "2026-05-10", "標準B", "予定"],
        ],
    )
    make_book(
        input_folder / "一覧_20260421.xlsx",
        [
            ["apr", "2026-04-10", "標準A", "完了"],
            # may が消えた → 5月の延期として記録
        ],
    )

    run(datetime.date(2026, 4, 25))  # 23日以降なので 4月と5月が対象

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    target_months = {row["対象月"] for row in rows}
    assert "2026-04" in target_months
    assert "2026-05" in target_months


def test_range_floor_returns_first_day_of_previous_month() -> None:
    """今日が5月15日 → 前月の1日（4月1日）に切り上げる。"""
    assert _range_floor(datetime.date(2026, 5, 15)) == datetime.date(2026, 4, 1)


def test_range_floor_wraps_to_previous_year_in_january() -> None:
    """1月実行時、前月の1日は前年12月1日に戻る。"""
    assert _range_floor(datetime.date(2026, 1, 10)) == datetime.date(2025, 12, 1)
    assert _range_floor(datetime.date(2026, 1, 1)) == datetime.date(2025, 12, 1)


def test_run_does_not_open_files_older_than_range_floor(
    tmp_path: Path, make_book, setup_run
) -> None:
    """前月の初日より古いファイルのうち、比較相手の直前1件だけは例外的に開かれる。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 1月15日実行 → 下限は前年12月1日
    # 11/30 は 12/01 の比較相手として開かれてよい
    make_book(input_folder / "一覧_20251130.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    # 12/01 が範囲内最初 → 比較相手は 11/30（範囲外だが直前1件だけOK）
    make_book(input_folder / "一覧_20251201.xlsx", [["a", "2026-01-05", "標準A", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["a", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260112.xlsx", [["a", "2026-01-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 1, 15))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 11/30 は比較相手として開かれる（仕様で許される）
    assert "一覧_20251130.xlsx" in opened_names
    # 下限当日の 12/01 は範囲内
    assert "一覧_20251201.xlsx" in opened_names
    # 範囲内ファイルも開かれる
    assert "一覧_20260110.xlsx" in opened_names
    assert "一覧_20260112.xlsx" in opened_names


def test_run_skips_files_older_than_the_predecessor(
    tmp_path: Path, make_book, setup_run
) -> None:
    """下限より古いファイルのうち、比較相手の直前1件より更に古いものは開かれない。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 1月15日実行 → 下限は前年12月1日
    # 比較相手の候補が 11/25, 11/28, 11/30 と3つあるが、開かれるのは 11/30 だけ
    make_book(input_folder / "一覧_20251125.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251128.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251130.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251201.xlsx", [["a", "2026-01-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 1, 15))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 11/25 と 11/28 は開かれない（比較相手の直前1件より更に古い）
    assert "一覧_20251125.xlsx" not in opened_names
    assert "一覧_20251128.xlsx" not in opened_names
    # 直前の 11/30 だけが比較相手として開かれる
    assert "一覧_20251130.xlsx" in opened_names
    # 範囲内の 12/01 も開かれる
    assert "一覧_20251201.xlsx" in opened_names
    # 開かれたのは2件だけ
    assert len(opened) == 2


def test_run_reads_only_one_predecessor_outside_the_range(
    tmp_path: Path, make_book, setup_run
) -> None:
    """範囲外に古いファイルが複数あっても、比較相手として開くのは直前1件だけ。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 4月10日実行 → 下限は3月1日。範囲外（3/1より前）のファイルが複数ある
    # 3/1 の比較相手として開かれるのは範囲外のうち最も新しい 2/5 だけ
    make_book(input_folder / "一覧_20260130.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260131.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260205.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260301.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260302.xlsx", [["x", "2026-03-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 4, 10))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 範囲外で比較相手にもならない古いファイルは開かれない
    assert "一覧_20260130.xlsx" not in opened_names
    assert "一覧_20260131.xlsx" not in opened_names
    # 直前の 2/5 だけは比較相手として開かれる
    assert "一覧_20260205.xlsx" in opened_names
    # 範囲内ファイルも開かれる
    assert "一覧_20260301.xlsx" in opened_names
    assert "一覧_20260302.xlsx" in opened_names
    # 開かれたのは合計4件（比較相手1 + 範囲内2 + 範囲内3/1より後→ 3/1, 3/2）
    assert len(opened) == 3


def test_run_only_opens_one_predecessor_when_multiple_outside_exist(
    tmp_path: Path, make_book, setup_run
) -> None:
    """範囲外に古いファイルが2つ以上あっても、比較相手として開かれるのは直前の1つだけ。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 4月10日実行 → 下限は3月1日
    # 範囲外の古いファイルが複数：2/25, 2/27, 2/28（3/1 の比較相手として開かれるのは 2/28 だけ）
    make_book(input_folder / "一覧_20260225.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260227.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["x", "2026-03-05", "標準A", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260301.xlsx", [["x", "2026-03-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 4, 10))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 範囲外の 2/25 と 2/27 は開かれてはいけない
    assert "一覧_20260225.xlsx" not in opened_names
    assert "一覧_20260227.xlsx" not in opened_names
    # 直前の 2/28 だけが比較相手として開かれる
    assert "一覧_20260228.xlsx" in opened_names
    # 範囲内の 3/1 も開かれる
    assert "一覧_20260301.xlsx" in opened_names
    # 全体で2つしか開かれていない
    assert len(opened) == 2


def test_run_does_not_open_files_after_today(
    tmp_path: Path, make_book, setup_run
) -> None:
    """今日より後の日付のファイルは開かれない。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 4月10日実行 → 範囲は 3/1〜4/10
    # 範囲内最初（3/5）の比較相手として 2/28 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260305.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260408.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260410.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
    # 今日より後
    make_book(input_folder / "一覧_20260415.xlsx", [["a", "2026-04-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 4, 10))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 今日（4/10）以前は開かれてよいが、今日以降（4/15）は開かれない
    assert "一覧_20260415.xlsx" not in opened_names
    # 範囲内のファイルは開かれる
    assert "一覧_20260410.xlsx" in opened_names
    assert "一覧_20260408.xlsx" in opened_names
    assert "一覧_20260305.xlsx" in opened_names
    # 比較相手の 2/28 も開かれる（範囲外だが直前1件）
    assert "一覧_20260228.xlsx" in opened_names


def test_run_skips_when_no_files_in_range(
    tmp_path: Path, make_book, setup_run
) -> None:
    """範囲（前月の初日〜今日）内にファイルが無いとき、落ちずにログを出して終わる。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 4月10日実行 → 範囲は 3/1〜4/10。すべて範囲外のファイルしかない
    make_book(input_folder / "一覧_20260225.xlsx", [["a", "2026-02-20", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        # 落ちないことを確認（戻り値は output_path）
        result = run(datetime.date(2026, 4, 10))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    # ファイルは何も開かれない
    assert opened == []
    assert result == output_folder / OUTPUT_NAME


def _cell(
    rows: list[dict[str, object]],
    target_month: str,
    plan: str,
    status: str,
    date_header: str,
) -> str:
    row = _row(rows, target_month, plan, status)
    value = row.get(date_header, "")
    return str(value)


def _row(
    rows: list[dict[str, object]],
    target_month: str,
    plan: str,
    status: str,
) -> dict[str, object]:
    for row in rows:
        if row["対象月"] == target_month and row["種別"] == plan and row["判定"] == status:
            return row
    raise AssertionError(f"行が見つかりません: {target_month} {plan} {status}")