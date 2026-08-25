import datetime
import logging
from pathlib import Path
from unittest.mock import patch

import pytest
from comken.toolbox.csv import CSV

from src.run import OUTPUT_NAME, _range_floor, _target_months, run


def _config_text(
    input_folder: Path,
    output_folder: Path,
    *,
    rolling_window_days: int | None = None,
) -> str:
    """テスト用の config.ini 本文。

    ``rolling_window_days`` を指定すると ``[FILES] ROLLING_WINDOW_DAYS`` 行を追加する。
    ``None`` のときは追加しない（年次累積モード既定）。
    """
    rolling_line = (
        f"ROLLING_WINDOW_DAYS = {rolling_window_days}\n"
        if rolling_window_days is not None
        else ""
    )
    return f"""[FILES]
INPUT_FOLDER = {input_folder}
FILE_PATTERN = 一覧_*.xlsx
{rolling_line}[SOURCE]
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
def setup_run(config_for_tests, monkeypatch):
    """input/output フォルダを tmp に作り、config.ini を読み込む。

    ``run()`` は ``comken.core.today`` から現在日付を取る（テストのたびに違う
    「今日」を使う）ので、ここで固定値を差し込む。
    """

    def _setup(
        tmp_path: Path,
        *,
        kinds_value: str = "[完了, 予定]",
        today_date: datetime.date | None = None,
        rolling_window_days: int | None = None,
    ) -> tuple[Path, Path]:
        input_folder = tmp_path / "input"
        output_folder = tmp_path / "output"
        input_folder.mkdir()
        output_folder.mkdir()
        config_for_tests(_config_text(
            input_folder, output_folder, rolling_window_days=rolling_window_days,
        ).replace(
            "KINDS = [完了, 予定]", f"KINDS = {kinds_value}"
        ))
        if today_date is not None:
            # ``src.run.today`` を固定値で差し替える（テストのたびに違う「今日」を使う）。
            # ``run.py`` は ``from comken.core import today`` で束縛しているので、
            # ``comken.core.today`` だけ差し替えても src.run 側に届かない
            monkeypatch.setattr("src.run.today", lambda: today_date)
        return input_folder, output_folder

    return _setup


def test_target_months_uses_business_date_month() -> None:
    """業務日 = その業務日の属する月の文字列。1つの業務日に対して1つの対象月。"""
    assert _target_months(datetime.date(2026, 4, 10)) == ["2026-04"]
    assert _target_months(datetime.date(2026, 12, 25)) == ["2026-12"]


def test_target_months_rolling_mode_keeps_only_current_month_when_far_from_end() -> None:
    """ローリングモードでも、月の前半（残日数が窓幅以上）は当月だけ返す。

    4月の業務日 4/10・窓幅 7 → 残日数 = 20。20 < 7 は偽なので lookahead しない。
    """
    assert _target_months(
        datetime.date(2026, 4, 10), rolling_window_days=7,
    ) == ["2026-04"]
    # 12月→1月の年跨ぎも、月の前半なら当月だけ
    assert _target_months(
        datetime.date(2026, 12, 10), rolling_window_days=7,
    ) == ["2026-12"]


def test_target_months_rolling_mode_includes_next_month_near_month_end() -> None:
    """ローリングモードで月末近く（残日数が窓幅未満）なら翌月も対象月に加える。

    4/25・窓幅 7 → 残日数 = 5（= 4/30 - 4/25）。5 < 7 は真なので 4月 + 5月 を返す。
    """
    assert _target_months(
        datetime.date(2026, 4, 25), rolling_window_days=7,
    ) == ["2026-04", "2026-05"]
    # 月末当日も翌月を含める（残日数 = 0 < 7）
    assert _target_months(
        datetime.date(2026, 4, 30), rolling_window_days=7,
    ) == ["2026-04", "2026-05"]
    # 12月末→1月の年跨ぎ
    assert _target_months(
        datetime.date(2026, 12, 31), rolling_window_days=7,
    ) == ["2026-12", "2027-01"]


def test_target_months_rolling_mode_boundary_days_to_end_equals_window() -> None:
    """境界値: 残日数が窓幅 **ちょうど** のときは lookahead しない（< 7 の判定）。

    4/23・窓幅 7 → 残日数 = 7（= 4/30 - 4/23）。7 < 7 は偽なので当月のみ。
    """
    assert _target_months(
        datetime.date(2026, 4, 23), rolling_window_days=7,
    ) == ["2026-04"]


def test_target_months_rolling_mode_boundary_days_to_end_equals_window_minus_one() -> None:
    """境界値: 残日数が窓幅 - 1 のときは lookahead する。

    4/24・窓幅 7 → 残日数 = 6（= 4/30 - 4/24）。6 < 7 は真なので翌月も加える。
    """
    assert _target_months(
        datetime.date(2026, 4, 24), rolling_window_days=7,
    ) == ["2026-04", "2026-05"]


def test_target_months_rolling_mode_uses_calendar_month_end_not_business_days() -> None:
    """月末判定は暦月末（``calendar.monthrange``）。4月は30日、2月は28/29日で計算する。

    窓幅 31 のとき: 4月なら 4/1 は残日数 29 で「29 < 31」となり lookahead する。
    """
    # 4月1日・窓幅 31 → 残日数 29 < 31 で 4月 + 5月
    assert _target_months(
        datetime.date(2026, 4, 1), rolling_window_days=31,
    ) == ["2026-04", "2026-05"]
    # 2月1日（平年・28日）・窓幅 31 → 残日数 27 < 31 で 2月 + 3月
    assert _target_months(
        datetime.date(2026, 2, 1), rolling_window_days=31,
    ) == ["2026-02", "2026-03"]
    # 2月1日（閏年・29日）・窓幅 31 → 残日数 28 < 31 で 2月 + 3月
    assert _target_months(
        datetime.date(2024, 2, 1), rolling_window_days=31,
    ) == ["2024-02", "2024-03"]


def test_target_months_non_rolling_mode_ignores_lookahead() -> None:
    """年次累積モード（``rolling_window_days=None`` 既定）は lookahead しない。

    月末近くでも、当月だけ返す。既存の年次累積モードの挙動を完全保持。
    """
    assert _target_months(datetime.date(2026, 4, 25), rolling_window_days=None) == ["2026-04"]
    assert _target_months(datetime.date(2026, 12, 31), rolling_window_days=None) == ["2026-12"]


def test_run_starts_from_oldest_file_when_no_csv_exists(
    tmp_path: Path, make_book, setup_run
) -> None:
    """初回実行: 4 行（当月×{積み上げ/延期}, 来月×{積み上げ/延期}）構成で集計される。

    ローリングモードの設定は無し（年次累積モード既定）なので、「来月」行は構造上
    存在するが実データは入らない。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
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

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 行は常に 4 行（当月/来月 × 積み上げ/延期）。種別（標準/上位）の行は無い
    assert len(rows) == 4
    labels = [_row_label(r) for r in rows]
    assert labels == [
        ("当月", "積み上げ"),
        ("当月", "延期"),
        ("来月", "積み上げ"),
        ("来月", "延期"),
    ]

    # 業務日 4/21 (一覧_20260421.xlsx が無いので 一覧_20260420.xlsx vs 一覧_20260422.xlsx の比較)
    # → add が 4/22 ファイルに新しく出現 → 業務日 4/21 に積み上げ 1
    assert _cell(rows, "当月", "積み上げ", "2026-04-21") == "1"
    assert _cell(rows, "当月", "延期", "2026-04-21") == "0"
    # 業務日 4/22 (一覧_20260422.xlsx vs 一覧_20260423.xlsx)
    # → add が 4/23 ファイルから消えた → 業務日 4/22 に延期 1
    assert _cell(rows, "当月", "積み上げ", "2026-04-22") == "0"
    assert _cell(rows, "当月", "延期", "2026-04-22") == "1"


def test_run_writes_blank_for_business_days_without_files(
    tmp_path: Path, make_book, setup_run
) -> None:
    """業務日基準で、ファイルが無い業務日は空セル・件数が 0 の日は "0" になる。

    一覧_YYYYMMDD.xlsx は前日終了時点のデータなので、列は業務日（= ファイル日付 - 1日）で並ぶ。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    # 2/28 の比較相手として 2025/12/20 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    # 業務日 4/22 のためのファイル（=ファイル日付 4/23）はあるが、業務日 4/21 のための
    # ファイル（=ファイル日付 4/22）が無い
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    # 業務日 4/20 列: 一覧_20260420.xlsx vs 一覧_20260421.xlsx → 差分なし → "0"
    assert _cell(rows, "当月", "積み上げ", "2026-04-20") == "0"
    # 業務日 4/21 列: 一覧_20260422.xlsx が無いので比較できず → 空セル
    assert _cell(rows, "当月", "積み上げ", "2026-04-21") == ""
    # 業務日 4/22 列: 一覧_20260423.xlsx vs 一覧_20260422.xlsx（無しのため直前の 4/21 比較）
    # → 差分なし → "0"
    assert _cell(rows, "当月", "積み上げ", "2026-04-22") == "0"
    # 業務日 4/23 列: 一覧_20260424.xlsx が無いので列が無い → 空セル
    assert _cell(rows, "当月", "積み上げ", "2026-04-23") == ""


def test_run_continues_from_last_csv_date_and_keeps_existing_columns(
    tmp_path: Path, make_book, setup_run
) -> None:
    """増分実行でも行は常に 4 行固定。"""
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    # 1回目: 2025/12/20（2/28 の比較相手）, 2/28, 4/20, 4/21 を実行
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run()
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    # 業務日 4/20 列 = 一覧_20260421.xlsx vs 一覧_20260420.xlsx の比較結果
    first_4_20 = _cell(first_rows, "当月", "積み上げ", "2026-04-20")

    # 2回目: 4/22, 4/23 のファイルを追加して再実行
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run()

    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())
    # 行は常に 4 行固定
    assert len(second_rows) == 4
    # 業務日 4/20 の値が変わらない
    assert _cell(second_rows, "当月", "積み上げ", "2026-04-20") == first_4_20
    # 業務日 4/21, 4/22 が追加される（ファイル日付 4/22, 4/23 でそれぞれ作られる）
    assert "2026-04-21" in second_rows[0]
    assert "2026-04-22" in second_rows[0]
    # 業務日 4/23 は対応するファイル（ファイル日付 4/24）が無いので列が無い
    assert "2026-04-23" not in second_rows[0]
    # 同じキー (当月, 積み上げ) の行が1つだけ（マージされている）
    matching = [r for r in second_rows
                if r["対象月"] == "当月" and r["判定"] == "積み上げ"]
    assert len(matching) == 1


def test_run_row_has_values_only_for_target_month_dates(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ラベル ``当月`` 行は、当月だった業務日の列にだけ値を持つ。"""
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 25)
    )
    # 業務日 1/22 まで: 1月の対象月のみ
    make_book(input_folder / "一覧_20251220.xlsx", [["x", "2026-01-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20251231.xlsx", [["x", "2026-01-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260122.xlsx", [["x", "2026-01-10", "標準A", "完了"]])
    # 業務日 1/22 = ファイル 1/23。x が消える
    make_book(input_folder / "一覧_20260123.xlsx", [])

    # 業務日 1/23 = ファイル 1/24。jan が増える（1月の対象月のみ）
    make_book(
        input_folder / "一覧_20260124.xlsx",
        [
            ["jan", "2026-01-15", "標準A", "完了"],
        ],
    )

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 業務日 1/22: x が消えた → 当月の行に「延期 1」
    assert _cell(rows, "当月", "延期", "2026-01-22") == "1"

    # 業務日 1/23: jan が増えた → 当月の行に「積み上げ 1」
    assert _cell(rows, "当月", "積み上げ", "2026-01-23") == "1"


def _row_label(row: dict[str, object]) -> tuple[str, str]:
    return (str(row["対象月"]), str(row["判定"]))


def test_run_reads_each_file_once_across_month_boundary(
    tmp_path: Path, make_book, setup_run
) -> None:
    """月をまたぐ業務日（1/31 と 2/1）で、一覧_20260201.xlsx が両方の計算に使われるが
    1回しか読まれない。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 2, 3)
    )
    # 1/30 ファイルの比較相手として 2025/12/20 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["x", "2026-01-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260130.xlsx", [["x", "2026-01-10", "標準A", "完了"]])
    # 月末のファイル（1/31 終了時点 → 業務日 1/30 で使う）
    make_book(input_folder / "一覧_20260131.xlsx", [["x", "2026-01-10", "標準A", "完了"]])
    # 月またぎ。一覧_20260201.xlsx は:
    #   - 業務日 1/31 (= 1/31 file vs 2/1 file) の current → 対象月 1月
    #   - 業務日 2/1  (= 2/1 file vs 2/2 file) の previous → 対象月 2月
    make_book(
        input_folder / "一覧_20260201.xlsx",
        [
            ["jan", "2026-01-15", "標準A", "完了"],
            ["feb", "2026-02-05", "標準B", "予定"],
        ],
    )
    # 2/2 ファイル（業務日 2/1 の current）
    make_book(
        input_folder / "一覧_20260202.xlsx",
        [
            ["jan", "2026-01-15", "標準A", "完了"],
            ["feb", "2026-02-05", "標準B", "予定"],
        ],
    )

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {"jan": __import__("src.source", fromlist=["Record"]).Record(
            "jan", datetime.date(2026, 1, 15), "標準"),
            "feb": __import__("src.source", fromlist=["Record"]).Record(
                "feb", datetime.date(2026, 2, 5), "標準"),
            "x": __import__("src.source", fromlist=["Record"]).Record(
                "x", datetime.date(2026, 1, 10), "標準"),
        }

    with patch("src.diff.read_records", side_effect=_capture):
        run()

    opened_names = [p.name for p in opened]
    # 一覧_20260201.xlsx は1回しか読まれない
    assert opened_names.count("一覧_20260201.xlsx") == 1
    # すべてのファイルが1回ずつ
    assert len(opened_names) == len(set(opened_names))


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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 15)
    )
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
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 15)
    )
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
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 15)
    )
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
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 15)
    )
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
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 10)
    )
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
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 10)
    )
    # 1月10日実行 → 範囲は 1/1〜1/10。範囲内のファイルが無く、前年の12月ファイルも
    # in_range が空なので比較相手としても読まれない
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2025-12-20", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        # 落ちないことを確認（戻り値は output_path）
        result = run()

    print("\n[opened-files] " + ", ".join(p.name for p in opened))
    # ファイルは何も開かれない
    assert opened == []
    assert result == output_folder / OUTPUT_NAME


def test_run_writes_csv_after_each_day(
    tmp_path: Path, make_book, setup_run
) -> None:
    """1日ぶん計算するたびに CSV が書き出される（N 件ごとに区切らない）。"""
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    save_calls: list[Path] = []
    original_write_csv = __import__("src.report", fromlist=["write_csv"]).write_csv

    def _spy_write_csv(
        path, by_row, dates, row_keys, target_dates_by_month=None,
    ):
        save_calls.append(path)
        # 本物 ``write_csv`` は ``(行数, 列数)`` を返すので、spy もその戻り値を
        # そのまま返す（``_make_day_saver`` 内の ``sizes.append`` が None で
        # 落ちないようにするため）
        return original_write_csv(
            path, by_row, dates, row_keys, target_dates_by_month,
        )

    with patch("src.run.write_csv", side_effect=_spy_write_csv):
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
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
                run()

    # 例外までに書き出した CSV にはそこまでの分の列が残っている
    assert csv_path.exists()
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "判定")]
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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
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
            run()

    csv_path = output_folder / OUTPUT_NAME
    assert csv_path.exists()
    with CSV(csv_path) as csv:
        first_rows = list(csv.read())
    first_headers = [h for h in first_rows[0].keys() if h not in ("対象月", "判定")]
    # 1回目の保存：業務日 4/20 列まで書き込まれている
    # （4/21 ファイル vs 4/20 ファイルの比較結果が業務日 4/20 列に入る）
    assert "2026-04-20" in first_headers
    # 業務日 4/21, 4/22 列はまだ無い
    assert "2026-04-21" not in first_headers
    assert "2026-04-22" not in first_headers
    first_4_20 = _cell(first_rows, "当月", "積み上げ", "2026-04-20")

    # 2回目: そのまま再開（モック解除）
    run()

    with CSV(csv_path) as csv:
        second_rows = list(csv.read())
    second_headers = [h for h in second_rows[0].keys() if h not in ("対象月", "判定")]
    # 業務日 4/21, 4/22 列まで書き込まれている（一覧_20260422.xlsx と 一覧_20260423.xlsx の比較）
    assert "2026-04-21" in second_headers
    assert "2026-04-22" in second_headers
    # 業務日 4/20 の値は変わらない（4/19 列は前回 save の最終列）
    second_4_20 = _cell(second_rows, "当月", "積み上げ", "2026-04-20")
    assert second_4_20 == first_4_20


def test_run_offsets_column_by_one_business_day(
    tmp_path: Path, make_book, setup_run
) -> None:
    """一覧_20260824.xlsx と 一覧_20260825.xlsx の差分は「8/24 に動いたぶん」。

    ファイル名は 8/24 と 8/25 だが、列は **業務日** の 8/24 (= 8/25 - 1日) に記録される。
    8/25 の列に書かれると勘違いしやすいので日付で固定する。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 8, 26)
    )
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

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())

    # 業務日 8/24 の列が存在し、id_b が「積み上げ」として記録される
    assert "2026-08-24" in headers
    assert "2026-08-25" not in headers
    # 一覧_20260824.xlsx と 一覧_20260825.xlsx の差分は業務日 8/24 の列に入る
    assert _cell(rows, "当月", "積み上げ", "2026-08-24") == "1"
    assert _cell(rows, "当月", "延期", "2026-08-24") == "0"


def test_run_advances_one_file_in_incremental_run(
    tmp_path: Path, make_book, setup_run
) -> None:
    """既存 CSV の最後の列が業務日 8/24 のとき、次はファイル日付 8/26 から読む。

    業務日 8/25 の列を作るには「ファイル日付 8/26」（= 8/25 終了時点）が必要。
    ファイル日付 8/25 から作り直さず、ファイル日付 8/26 以降だけ読むことを
    実際にファイルを読んで確認する（off-by-one 防止）。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 8, 26)
    )
    # 8/23 ファイルの比較相手として 2025/12/20 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260823.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260824.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260825.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])
    run()
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    last_date_header = max(
        h for h in first_rows[0].keys() if h not in ("対象月", "判定")
    )
    assert last_date_header == "2026-08-24"

    # 翌日分を追加して再実行
    make_book(input_folder / "一覧_20260826.xlsx", [["id_a", "2026-08-23", "標準A", "完了"]])

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {}

    with patch("src.diff.read_records", side_effect=_capture):
        run()

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
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 9, 2)
    )
    # 8/28 ファイルの比較相手として 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260828.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260829.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260830.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260831.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260901.xlsx", [["id_a", "2026-08-28", "標準A", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "判定")]

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
    label: str,
    status: str,
    date_header: str,
) -> str:
    row = _row(rows, label, status)
    value = row.get(date_header, "")
    return str(value)


def _row(
    rows: list[dict[str, object]],
    label: str,
    status: str,
) -> dict[str, object]:
    for row in rows:
        if row["対象月"] == label and row["判定"] == status:
            return row
    raise AssertionError(f"行が見つかりません: {label} {status}")


# ====================================================================
# ローリングモードのテスト（[FILES] ROLLING_WINDOW_DAYS）
# --------------------------------------------------------------------
# 既存テスト群は config.ini に ROLLING_WINDOW_DAYS が無い前提（非ローリング/
# 年次累積）で書かれており、すべて ROLLING_WINDOW_DAYS 未設定でもパスする
# ＝「キーが無いときは現状の年次累積モードのまま」の回帰確認になっている。
# ここではローリングモード固有の挙動を追加で検証する。
# ====================================================================


def test_range_floor_returns_rolling_window_floor() -> None:
    """``_range_floor(today, rolling_window_days=N)`` は今日から ``N-1`` 日前を返す。

    N=1 なら今日だけ、N=7 なら今日から6日前まで。月またぎも正しく動く。
    """
    today = datetime.date(2026, 4, 25)
    assert _range_floor(today, rolling_window_days=1) == datetime.date(2026, 4, 25)
    assert _range_floor(today, rolling_window_days=7) == datetime.date(2026, 4, 19)
    assert _range_floor(today, rolling_window_days=30) == datetime.date(2026, 3, 27)
    # 1月をまたぐ計算も単純な暦日差で良い（年末調整は要件外）
    year_end = datetime.date(2026, 1, 3)
    assert _range_floor(year_end, rolling_window_days=7) == datetime.date(2025, 12, 28)


def test_range_floor_returning_none_keeps_yearly_cumulative_behavior() -> None:
    """``rolling_window_days=None``（既定）のときは今年の1月1日（既存の挙動）。"""
    # 既存のテストがすべてこのパスを通る = 既存の年次累積モードと完全同一。
    assert _range_floor(datetime.date(2026, 5, 15), rolling_window_days=None) == datetime.date(
        2026, 1, 1
    )
    assert _range_floor(datetime.date(2026, 1, 5), rolling_window_days=None) == datetime.date(
        2026, 1, 1
    )


def test_rolling_mode_first_run_only_writes_columns_within_window(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ローリングモード初回実行は、窓の外（より古い業務日）の列を作らない。

    今日=4/25・窓 7日 → 業務日 4/19 以降が窓。4/19 ファイル日付より古い
    ファイルは比較相手として1ファイルだけ読まれるが、業務日 4/19 未満の
    列は保存されない。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25), rolling_window_days=7,
    )
    # 4/18（業務日 4/17）は窓の下限 4/19 より古いが、4/19 の比較相手として読まれる
    make_book(input_folder / "一覧_20260418.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260419.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "判定")]
    # 窓の下限 = 業務日 4/19。今日が 4/25 で、今日ファイル（= 4/26）は無いので
    # 比較が成立するのは 4/19・4/20・4/21。保存時は start_date から last_compared
    # までの range を必ず列に出す
    assert "2026-04-19" in date_headers
    assert "2026-04-20" in date_headers
    assert "2026-04-21" in date_headers
    # 業務日 4/18 未満の列は作られない（窓の外）
    assert "2026-04-18" not in date_headers
    assert "2026-04-17" not in date_headers
    # 行は常に 4 行固定
    assert len(rows) == 4


def test_rolling_mode_year_cumulative_keeps_next_month_row_even_when_empty(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ローリングモードでない年次累積実行では「来月」行も常に存在するが、月の前半は
    lookahead が起きないので実データは入らない（構造上は空のまま）。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 15)
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260413.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260414.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260415.xlsx", [["a", "2026-04-10", "標準A", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    labels = [(r["対象月"], r["判定"]) for r in rows]
    # 4 行固定（当月/来月 × 積み上げ/延期）
    assert labels == [
        ("当月", "積み上げ"),
        ("当月", "延期"),
        ("来月", "積み上げ"),
        ("来月", "延期"),
    ]
    # 年次累積モードで月の前半 → 「来月」行は構造上あるが、月初の業務日しか
    # 比較していないので実データなし（対象月はすべて 4月のまま = 「当月」）
    next_added = _row(rows, "来月", "積み上げ")
    next_postponed = _row(rows, "来月", "延期")
    # データが無い業務日列は空セル、比較対象で対象月だったら "0"。
    # 月の途中で全て 4 月以外の業務日がなければ「来月」の対象月は発生しない
    # → 対象月の列は "0" / "" が混在で、件数セルとして 0 や "" が入る
    # 「来月」行（積み上げ・延期の双方）が **4 月の業務日に件数を持っていない**ことを確認
    for row_dict in (next_added, next_postponed):
        for header, value in row_dict.items():
            if header == "対象月" or header == "判定":
                continue
            assert value == "0" or value == ""


def test_rolling_mode_writes_to_next_month_row_near_month_end(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ローリングモード + 月末近くで、``来月`` 行に件数が入る。

    今日=4/25・窓 7。4/25 は残日数 5（< 7）なので lookahead 発火 → 4/25 vs 4/26
    の比較結果が「当月（4月）」と「来月（5月）」の両方に入る。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25), rolling_window_days=7,
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260419.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(
        input_folder / "一覧_20260425.xlsx",
        [
            ["a", "2026-04-10", "標準A", "完了"],
            ["new_apr", "2026-04-15", "標準A", "完了"],
            ["new_may", "2026-05-02", "標準A", "完了"],
        ],
    )

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 4 ファイル → 比較対象は 3 回目以降（4/19, 4/24）。4/24 で「来月 = 5月」も対象
    # 当月 / 来月の両方で 4/24 列を見る
    # 当月の行: 4月案件 new_apr が増えた → 当月・積み上げ = 1
    # 来月の行: 5月案件 new_may が増えた → 来月・積み上げ = 1
    assert _cell(rows, "当月", "積み上げ", "2026-04-24") == "1"
    assert _cell(rows, "来月", "積み上げ", "2026-04-24") == "1"
    assert _cell(rows, "当月", "延期", "2026-04-24") == "0"
    assert _cell(rows, "来月", "延期", "2026-04-24") == "0"


def test_non_rolling_mode_key_unchanged_in_existing_tests(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ROLLING_WINDOW_DAYS キーが無いとき、1月1日からの年次累積で動く（回帰）。

    既存のテスト群は config.ini に ROLLING_WINDOW_DAYS を含めない前提で、
    すべてこのパスを通る。
    """
    # setup_run を rolling_window_days=None で呼ぶ = キー無し
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25),
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260102.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "完了"]])
    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "判定")]
    # 年次累積モード: 今年の 1/1 〜 今日（4/25）が対象
    assert any(h.startswith("2026-01") for h in date_headers), (
        f"1月の列が存在すること（年次累積）: {date_headers}"
    )
    # 4月の列も出ている
    assert any(h.startswith("2026-04") for h in date_headers), (
        f"4月の列が存在すること（年次累積）: {date_headers}"
    )
    # 行は常に 4 行固定
    assert len(rows) == 4
