import datetime
import logging
from pathlib import Path
from unittest.mock import patch

import pytest
from comken.toolbox.csv import CSV

from src.run import CONDITIONS_NAME, OUTPUT_NAME, _range_floor, _target_months, run


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
    # 2/28 の比較相手として 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["stay", "2026-04-10", "標準A", "完了"]])
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
    # 業務日 4/21 (一覧_20260421.xlsx が無いので 一覧_20260420.xlsx vs 一覧_20260422.xlsx の比較)
    # → add が 4/22 ファイルに新しく出現 → 業務日 4/21 に積み上げ 1
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-21") == "1"
    assert _cell(rows, "2026-04", "標準", "延期", "2026-04-21") == "0"
    # 業務日 4/22 (一覧_20260422.xlsx vs 一覧_20260423.xlsx)
    # → add が 4/23 ファイルから消えた → 業務日 4/22 に延期 1
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-22") == "0"
    assert _cell(rows, "2026-04", "標準", "延期", "2026-04-22") == "1"


def test_run_writes_blank_for_business_days_without_files(
    tmp_path: Path, make_book, setup_run
) -> None:
    """業務日基準で、ファイルが無い業務日は空セル・件数が 0 の日は "0" になる。

    一覧_YYYYMMDD.xlsx は前日終了時点のデータなので、列は業務日（= ファイル日付 - 1日）で並ぶ。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 2/28 の比較相手として 2025/12/20 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    # 業務日 4/22 のためのファイル（=ファイル日付 4/23）はあるが、業務日 4/21 のための
    # ファイル（=ファイル日付 4/22）が無い
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    run(datetime.date(2026, 4, 25))

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    # 業務日 4/20 列: 一覧_20260420.xlsx vs 一覧_20260421.xlsx → 差分なし → "0"
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-20") == "0"
    # 業務日 4/21 列: 一覧_20260422.xlsx が無いので比較できず → 空セル
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-21") == ""
    # 業務日 4/22 列: 一覧_20260423.xlsx vs 一覧_20260422.xlsx（無しのため直前の 4/21 比較）
    # → 差分なし → "0"
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-22") == "0"
    # 業務日 4/23 列: 一覧_20260424.xlsx が無いので列が無い → 空セル
    assert _cell(rows, "2026-04", "標準", "積み上げ", "2026-04-23") == ""


def test_run_continues_from_last_csv_date_and_keeps_existing_columns(
    tmp_path: Path, make_book, setup_run
) -> None:
    input_folder, output_folder = setup_run(tmp_path)
    # 1回目: 2025/12/20（2/28 の比較相手）, 2/28, 4/20, 4/21 を実行
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run(datetime.date(2026, 4, 25))
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    # 業務日 4/20 列 = 一覧_20260421.xlsx vs 一覧_20260420.xlsx の比較結果
    first_4_20 = _cell(first_rows, "2026-04", "標準", "積み上げ", "2026-04-20")

    # 2回目: 4/22, 4/23 のファイルを追加して再実行
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run(datetime.date(2026, 4, 25))

    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())
    # 業務日 4/20 の値が変わらない
    assert _cell(second_rows, "2026-04", "標準", "積み上げ", "2026-04-20") == first_4_20
    # 業務日 4/21, 4/22 が追加される（ファイル日付 4/22, 4/23 でそれぞれ作られる）
    assert "2026-04-21" in second_rows[0]
    assert "2026-04-22" in second_rows[0]
    # 業務日 4/23 は対応するファイル（ファイル日付 4/24）が無いので列が無い
    assert "2026-04-23" not in second_rows[0]
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

    # 1回目: 全ファイルを実行（KINDS = [完了, 予定]）。2025/12/20 は 2/28 の比較相手
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
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
    # 4月のとき 4月のみが対象。比較相手用に 2025/12/20 と 2/28 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
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
    run(datetime.date(2026, 5, 25))  # 5月25日 → 5月と6月が対象。4/22〜4/23 が下限(4/1)より後

    # 4月の行 が残っていること、5月と6月の行が新たに作られていることを確認
    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())

    # 4月の行が残る
    assert any(row["対象月"] == "2026-04" for row in second_rows)
    # 5月と6月の行が新たに作られている
    assert any(row["対象月"] == "2026-05" for row in second_rows)
    assert any(row["対象月"] == "2026-06" for row in second_rows)
    # 4月の行の業務日 4/22 列は空（4月はもう対象月でないため、ループで処理されない）。
    # 業務日 4/23 列は対応するファイル（ファイル日付 4/24）が無いので存在しない。
    row_apr_added = _row(second_rows, "2026-04", "標準", "積み上げ")
    assert row_apr_added["2026-04-22"] == ""
    assert "2026-04-23" not in second_rows[0]


def test_run_targets_two_months_from_the_23rd(
    tmp_path: Path, make_book, setup_run
) -> None:
    input_folder, output_folder = setup_run(tmp_path)
    # 範囲内最初（2/28）の比較相手として 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
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


def test_range_floor_returns_first_day_of_current_year() -> None:
    """今日がいつであっても、その年の1月1日（前月の1日ではない）。"""
    assert _range_floor(datetime.date(2026, 5, 15)) == datetime.date(2026, 1, 1)
    assert _range_floor(datetime.date(2026, 8, 25)) == datetime.date(2026, 1, 1)


def test_range_floor_does_not_wrap_to_previous_year_in_january() -> None:
    """1月実行時も、同じ年の1月1日（前年の12月1日ではない）。"""
    assert _range_floor(datetime.date(2026, 1, 1)) == datetime.date(2026, 1, 1)
    assert _range_floor(datetime.date(2026, 1, 5)) == datetime.date(2026, 1, 1)
    assert _range_floor(datetime.date(2026, 1, 10)) == datetime.date(2026, 1, 1)


def test_run_does_not_open_files_older_than_range_floor(
    tmp_path: Path, make_book, setup_run
) -> None:
    """今年の1月1日より古いファイルのうち、比較相手の直前1件だけは例外的に開かれる。

    1月実行時、前年12月のファイルは範囲外だが、比較相手として最も新しい1件だけが読まれる。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 1月15日実行 → 下限は今年の1月1日
    # 前年12月のファイルのうち、最も新しい 12/20 が 1/10 の比較相手として開かれる
    make_book(input_folder / "一覧_20251201.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251205.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-01-05", "標準A", "完了"]])
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
    # 1月に実行したとき、前年12月のファイルは比較相手として1件だけ読まれる
    assert "一覧_20251220.xlsx" in opened_names
    # 比較相手より更に古い前年のファイルは開かれない
    assert "一覧_20251201.xlsx" not in opened_names
    assert "一覧_20251205.xlsx" not in opened_names
    # 範囲内のファイルも開かれる
    assert "一覧_20260110.xlsx" in opened_names
    assert "一覧_20260112.xlsx" in opened_names
    # 比較相手 1 件 + 範囲内 2 件 = 3 件
    assert len(opened) == 3


def test_run_skips_files_older_than_the_predecessor(
    tmp_path: Path, make_book, setup_run
) -> None:
    """下限より古いファイルのうち、比較相手の直前1件より更に古いものは開かれない。

    1月実行時、前年12月のファイルが複数あっても、比較相手として読まれるのは最も新しい1件だけ。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 1月15日実行 → 下限は今年の1月1日
    # 比較相手の候補が 12/15, 12/18, 12/20 と3つあるが、開かれるのは 12/20 だけ
    make_book(input_folder / "一覧_20251215.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251218.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251220.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["a", "2026-01-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 1, 15))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 12/15 と 12/18 は開かれない（比較相手の直前1件より更に古い）
    assert "一覧_20251215.xlsx" not in opened_names
    assert "一覧_20251218.xlsx" not in opened_names
    # 直前の 12/20 だけが比較相手として開かれる
    assert "一覧_20251220.xlsx" in opened_names
    # 範囲内の 1/10 も開かれる
    assert "一覧_20260110.xlsx" in opened_names
    # 開かれたのは2件だけ（比較相手 1 + 範囲内 1）
    assert len(opened) == 2


def test_run_reads_only_one_predecessor_outside_the_range(
    tmp_path: Path, make_book, setup_run
) -> None:
    """範囲外に前年のファイルが複数あっても、比較相手として開くのは直前1件だけ。

    1月実行時、前年12月のファイルが複数あっても、最も新しい1件だけが比較相手として読まれる。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 1月15日実行 → 下限は今年の1月1日
    # 前年12月のファイルが複数：12/01, 12/05, 12/20
    # 1/10 の比較相手として開かれるのは最も新しい 12/20 だけ
    make_book(input_folder / "一覧_20251201.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251205.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251220.xlsx", [["x", "2026-01-05", "標準A", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["x", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260112.xlsx", [["x", "2026-01-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 1, 15))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 前年12月のうち比較相手にならない古いファイルは開かれない
    assert "一覧_20251201.xlsx" not in opened_names
    assert "一覧_20251205.xlsx" not in opened_names
    # 直前の 12/20 だけが比較相手として開かれる
    assert "一覧_20251220.xlsx" in opened_names
    # 範囲内ファイルも開かれる
    assert "一覧_20260110.xlsx" in opened_names
    assert "一覧_20260112.xlsx" in opened_names
    # 比較相手 1 件 + 範囲内 2 件 = 3 件
    assert len(opened) == 3


def test_run_only_opens_one_predecessor_when_multiple_outside_exist(
    tmp_path: Path, make_book, setup_run
) -> None:
    """範囲外の前年ファイルが2つ以上あっても、比較相手として開かれるのは直前の1つだけ。

    1月実行時、前年12月のファイルが3つあっても、比較相手として読まれるのは最も新しい1件だけ。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 1月15日実行 → 下限は今年の1月1日
    # 前年12月のファイルが3つ：12/15, 12/18, 12/20
    # 1/10 の比較相手として開かれるのは最も新しい 12/20 だけ
    make_book(input_folder / "一覧_20251215.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251218.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251220.xlsx", [["old", "2026-01-05", "標準A", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["a", "2026-01-05", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 1, 15))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    opened_names = [p.name for p in opened]
    # 前年12月のうち 12/15 と 12/18 は開かれてはいけない
    assert "一覧_20251215.xlsx" not in opened_names
    assert "一覧_20251218.xlsx" not in opened_names
    # 直前の 12/20 だけが比較相手として開かれる
    assert "一覧_20251220.xlsx" in opened_names
    # 範囲内の 1/10 も開かれる
    assert "一覧_20260110.xlsx" in opened_names
    # 比較相手 1 + 範囲内 1 = 2 件
    assert len(opened) == 2


def test_run_does_not_open_files_after_today(
    tmp_path: Path, make_book, setup_run
) -> None:
    """今日より後の日付のファイルは開かれない。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 4月10日実行 → 範囲は 1/1〜4/10
    # 比較相手用に 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260201.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260405.xlsx", [["a", "2026-04-05", "標準A", "完了"]])
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
    assert "一覧_20260405.xlsx" in opened_names
    assert "一覧_20260201.xlsx" in opened_names
    # 比較相手の 2025/12/20 も開かれる（範囲外だが直前1件）
    assert "一覧_20251220.xlsx" in opened_names


def test_run_skips_when_no_files_in_range(
    tmp_path: Path, make_book, setup_run
) -> None:
    """範囲（今年の1月1日〜今日）内にファイルが無いとき、落ちずにログを出して終わる。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 1月10日実行 → 範囲は 1/1〜1/10。範囲内のファイルが無く、前年の12月ファイルも
    # in_range が空なので比較相手としても読まれない
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2025-12-20", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        # 落ちないことを確認（戻り値は output_path）
        result = run(datetime.date(2026, 1, 10))

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    # ファイルは何も開かれない
    assert opened == []
    assert result == output_folder / OUTPUT_NAME


def test_run_writes_csv_after_each_day(
    tmp_path: Path, make_book, setup_run
) -> None:
    """1日ぶん計算するたびに CSV が書き出される（N 件ごとに区切らない）。"""
    input_folder, output_folder = setup_run(tmp_path)
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    save_calls: list[Path] = []
    original_write_csv = __import__("src.report", fromlist=["write_csv"]).write_csv

    def _spy_write_csv(path, by_row, dates, row_keys):
        save_calls.append(path)
        original_write_csv(path, by_row, dates, row_keys)

    with patch("src.run.write_csv", side_effect=_spy_write_csv):
        run(datetime.date(2026, 4, 25))

    # targets = [12/20(predecessor), 2/28, 4/20, 4/21, 4/22, 4/23] のうち、
    # 比較が起きるのは2回目以降なので 5 回（2/28, 4/20, 4/21, 4/22, 4/23）。
    assert len(save_calls) == 5
    assert all(p == output_folder / OUTPUT_NAME for p in save_calls)


def test_run_keeps_partial_progress_when_read_fails_midway(
    tmp_path: Path, make_book, setup_run, caplog
) -> None:
    """読み込みが途中で失敗しても、それまでに書き出した分は CSV に残っている。

    列見出しは業務日なので、業務日 4/21（一覧_20260422.xlsx 由来）は保存済み、
    業務日 4/22（一覧_20260423.xlsx 由来）は保存されていない。
    """
    input_folder, output_folder = setup_run(tmp_path)
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    real_reader = __import__("src.diff", fromlist=["read_records"]).read_records
    call_count = {"n": 0}

    def _fail_on_sixth(path, *args, **kwargs):
        call_count["n"] += 1
        # 6回目の読み込み（=4/23）で例外。5回目までに4回保存されているはず。
        if call_count["n"] == 6:
            raise RuntimeError("simulated read failure")
        return real_reader(path, *args, **kwargs)

    csv_path = output_folder / OUTPUT_NAME
    with caplog.at_level(logging.INFO):
        with pytest.raises(RuntimeError):
            with patch("src.diff.read_records", side_effect=_fail_on_sixth):
                run(datetime.date(2026, 4, 25))

    # 例外までに書き出した CSV にはそこまでの分の列が残っている
    assert csv_path.exists()
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "種別", "判定")]
    # 業務日 4/21 (5回目の保存) はあり、業務日 4/22 はまだ無い
    assert "2026-04-21" in date_headers
    assert "2026-04-22" not in date_headers


def test_run_resumes_from_last_csv_date_after_crash(
    tmp_path: Path, make_book, setup_run
) -> None:
    """途中で落ちても、もう一度実行すると残っている最後の業務日の次の業務日から再開する。

    列見出しは業務日なので、4/22（業務日）の列は「4/22 vs 4/23」 = 一覧_20260423.xlsx で
    比較した結果。途中で落ちたときにどこまで書き込まれているかは、業務日で確認する。
    """
    input_folder, output_folder = setup_run(tmp_path)
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    real_reader = __import__("src.diff", fromlist=["read_records"]).read_records
    call_count = {"n": 0}

    def _fail_on_fifth(path, *args, **kwargs):
        call_count["n"] += 1
        # 5回目（=4/22 の read）で例外。3回保存されているはず
        # （2/28, 4/20, 4/21 の比較まで）。業務日 4/19 列まで保存。
        if call_count["n"] == 5:
            raise RuntimeError("simulated crash")
        return real_reader(path, *args, **kwargs)

    # 1回目: 5ファイル目でクラッシュ
    with pytest.raises(RuntimeError):
        with patch("src.diff.read_records", side_effect=_fail_on_fifth):
            run(datetime.date(2026, 4, 25))

    csv_path = output_folder / OUTPUT_NAME
    assert csv_path.exists()
    with CSV(csv_path) as csv:
        first_rows = list(csv.read())
    first_headers = [h for h in first_rows[0].keys() if h not in ("対象月", "種別", "判定")]
    # 1回目の保存：業務日 4/20 列まで書き込まれている
    # （4/21 ファイル vs 4/20 ファイルの比較結果が業務日 4/20 列に入る）
    assert "2026-04-20" in first_headers
    # 業務日 4/21, 4/22 列はまだ無い
    assert "2026-04-21" not in first_headers
    assert "2026-04-22" not in first_headers
    first_4_20 = _cell(first_rows, "2026-04", "標準", "積み上げ", "2026-04-20")

    # 2回目: そのまま再開（モック解除）
    run(datetime.date(2026, 4, 25))

    with CSV(csv_path) as csv:
        second_rows = list(csv.read())
    second_headers = [h for h in second_rows[0].keys() if h not in ("対象月", "種別", "判定")]
    # 業務日 4/21, 4/22 列まで書き込まれている（一覧_20260422.xlsx と 一覧_20260423.xlsx の比較）
    assert "2026-04-21" in second_headers
    assert "2026-04-22" in second_headers
    # 業務日 4/20 の値は変わらない（4/19 列は前回 save の最終列）
    second_4_20 = _cell(second_rows, "2026-04", "標準", "積み上げ", "2026-04-20")
    assert second_4_20 == first_4_20


def test_run_writes_conditions_file_on_first_save_even_if_run_crashes(
    tmp_path: Path, make_book, setup_run
) -> None:
    """1度目の保存の時点で条件ファイルを書いている（途中で落ちても次回が再開できる）。"""
    input_folder, output_folder = setup_run(tmp_path)
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    real_reader = __import__("src.diff", fromlist=["read_records"]).read_records
    call_count = {"n": 0}

    def _fail_on_fifth(path, *args, **kwargs):
        call_count["n"] += 1
        if call_count["n"] == 5:
            raise RuntimeError("simulated crash")
        return real_reader(path, *args, **kwargs)

    conditions_path = output_folder / CONDITIONS_NAME
    with pytest.raises(RuntimeError):
        with patch("src.diff.read_records", side_effect=_fail_on_fifth):
            run(datetime.date(2026, 4, 25))

    # 条件ファイルが書かれている（=初回保存より後に落ちた）
    assert conditions_path.exists()
    assert "KINDS" in conditions_path.read_text(encoding="utf-8")


def test_run_does_not_full_rebuild_when_conditions_change_crash_in_middle(
    tmp_path: Path, make_book, config_for_tests, caplog
) -> None:
    """条件変更による作り直しの途中で落ちても、次回は作り直しにならず続きから再開する。"""
    input_folder = tmp_path / "input"
    output_folder = tmp_path / "output"
    input_folder.mkdir()
    output_folder.mkdir()
    config_for_tests(_config_text(input_folder, output_folder))

    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    # 条件を変えてからクラッシュさせる（作り直しの途中で落ちる）
    config_for_tests(_config_text(input_folder, output_folder).replace(
        "KINDS = [完了, 予定]", "KINDS = [完了]"
    ))
    real_reader = __import__("src.diff", fromlist=["read_records"]).read_records
    call_count = {"n": 0}

    def _fail_on_fifth(path, *args, **kwargs):
        call_count["n"] += 1
        if call_count["n"] == 5:
            raise RuntimeError("simulated rebuild crash")
        return real_reader(path, *args, **kwargs)

    with pytest.raises(RuntimeError):
        with patch("src.diff.read_records", side_effect=_fail_on_fifth):
            run(datetime.date(2026, 4, 25))

    csv_path = output_folder / OUTPUT_NAME
    conditions_path = output_folder / CONDITIONS_NAME
    assert csv_path.exists()
    assert conditions_path.exists()

    # 再実行：条件は変えず、増分計算として動く。
    # 「実行モード: 増分」が出ること（＝全期間を作り直していない）で確認する。
    caplog.clear()
    with caplog.at_level(logging.INFO):
        run(datetime.date(2026, 4, 25))

    mode_logs = [r.message for r in caplog.records if r.message.startswith("実行モード:")]
    assert mode_logs == ["実行モード: 増分"]


def test_run_offsets_column_by_one_business_day(
    tmp_path: Path, make_book, setup_run
) -> None:
    """一覧_20260824.xlsx と 一覧_20260825.xlsx の差分は「8/24 に動いたぶん」。

    ファイル名は 8/24 と 8/25 だが、列は **業務日** の 8/24 (= 8/25 - 1日) に記録される。
    8/25 の列に書かれると勘違いしやすいので日付で固定する。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 8/24 ファイルの比較相手として 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["id_a", "2026-08-24", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260824.xlsx", [["id_a", "2026-08-24", "標準A", "完了"]])
    make_book(
        input_folder / "一覧_20260825.xlsx",
        [
            ["id_a", "2026-08-24", "標準A", "完了"],
            ["id_b", "2026-08-25", "標準B", "予定"],
        ],
    )

    run(datetime.date(2026, 8, 26))

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())

    # 業務日 8/24 の列が存在し、id_b が「積み上げ」として記録される
    assert "2026-08-24" in headers
    assert "2026-08-25" not in headers
    # 一覧_20260824.xlsx と 一覧_20260825.xlsx の差分は業務日 8/24 の列に入る
    assert _cell(rows, "2026-08", "標準", "積み上げ", "2026-08-24") == "1"
    assert _cell(rows, "2026-08", "標準", "延期", "2026-08-24") == "0"


def test_run_advances_one_file_in_incremental_run(
    tmp_path: Path, make_book, setup_run
) -> None:
    """既存 CSV の最後の列が業務日 8/24 のとき、次はファイル日付 8/26 から読む。

    業務日 8/25 の列を作るには「ファイル日付 8/26」（= 8/25 終了時点）が必要。
    ファイル日付 8/25 から作り直さず、ファイル日付 8/26 以降だけ読むことを
    実際にファイルを読んで確認する（off-by-one 防止）。
    """
    input_folder, output_folder = setup_run(tmp_path)
    # 8/23 ファイルの比較相手として 2025/12/20 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260823.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260824.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260825.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    run(datetime.date(2026, 8, 26))
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    last_date_header = max(
        h for h in first_rows[0].keys() if h not in ("対象月", "種別", "判定")
    )
    assert last_date_header == "2026-08-24"

    # 翌日分を追加して再実行
    make_book(input_folder / "一覧_20260826.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run(datetime.date(2026, 8, 26))

    opened_names = [p.name for p in opened]
    # 業務日 8/24 の列があるので、業務日 8/25 の列はファイル日付 8/26 から始める
    # → 一覧_20260824.xlsx は再計算のために開かれない（前回すでに集計済み）
    assert "一覧_20260824.xlsx" not in opened_names
    # 業務日 8/25 の列はファイル日付 8/26 で生成される
    assert "一覧_20260826.xlsx" in opened_names


def test_run_columns_span_month_boundary(
    tmp_path: Path, make_book, setup_run
) -> None:
    """月をまたいでも業務日列が連続する（8/29・8/30・8/31・9/1 が並ぶ）。"""
    input_folder, output_folder = setup_run(tmp_path)
    # 8/28 ファイルの比較相手として 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260828.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260829.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260830.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260831.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260901.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])

    run(datetime.date(2026, 9, 2))

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "種別", "判定")]

    # 業務日 8/28 (一覧_20260829.xlsx) 〜 業務日 8/31 (一覧_20260901.xlsx) が連続
    assert "2026-08-28" in date_headers
    assert "2026-08-29" in date_headers
    assert "2026-08-30" in date_headers
    assert "2026-08-31" in date_headers
    # 連続している（隣の列が必ず次の日）
    for prev, curr in zip(date_headers, date_headers[1:]):
        assert (datetime.date.fromisoformat(curr) - datetime.date.fromisoformat(prev)).days == 1


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