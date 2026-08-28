import datetime
import logging
from pathlib import Path
from unittest.mock import patch

import pytest
from comken.toolbox.csv import CSV

from src.report import ROW_LABELS
from src.run import OUTPUT_NAME, _range_floor, _target_months, run


def _config_text(
    input_folder: Path,
    output_folder: Path,
    *,
    rolling_window_days: int | None = None,
    plan_prefixes_value: str = "[標準, 上位]",
    crews_value: str = "[教育]",
    kinds_value: str = "[完了, 予定]",
) -> str:
    """テスト用の config.ini 本文。

    ``rolling_window_days`` を指定すると ``[FILES] ROLLING_WINDOW_DAYS`` 行を追加する。
    ``None`` のときは追加しない（年次累積モード既定）。
    ``plan_prefixes_value`` / ``crews_value`` / ``kinds_value`` で絞り込みを書き換えられる。
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
CREW_COLUMN = 施工作業班
KIND_COLUMN = 状態

[FILTER]
PLAN_PREFIXES = {plan_prefixes_value}
CREWS = {crews_value}
KINDS = {kinds_value}

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
        plan_prefixes_value: str = "[標準, 上位]",
        crews_value: str = "[教育]",
        today_date: datetime.date | None = None,
        rolling_window_days: int | None = None,
    ) -> tuple[Path, Path]:
        input_folder = tmp_path / "input"
        output_folder = tmp_path / "output"
        input_folder.mkdir()
        output_folder.mkdir()
        config_for_tests(_config_text(
            input_folder,
            output_folder,
            rolling_window_days=rolling_window_days,
            plan_prefixes_value=plan_prefixes_value,
            crews_value=crews_value,
            kinds_value=kinds_value,
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


def test_target_months_rolling_mode_always_includes_current_and_next_month() -> None:
    """ローリングモードでは、月初でも月末でも無条件で当月 + 翌月の2件を返す。

    旧仕様の「月末 N 日前からの lookahead」は廃止し、月のどの業務日でも
    ``[当月, 翌月]`` の2要素を返す。4月の月初から月末まで各日とも同じ結果になる。
    """
    for day in (1, 5, 10, 15, 23, 24, 25, 30):
        assert _target_months(
            datetime.date(2026, 4, day), rolling_window_days=7,
        ) == ["2026-04", "2026-05"]


def test_target_months_rolling_mode_handles_year_boundary_and_short_months() -> None:
    """ローリングモードの年またぎと、月末日数が 28/29 日の月でも翌月を返す。

    12月の業務日は翌年 1 月も対象月。2月の 28 日（平年）・29 日（閏年）も
    純粋な暦計算で翌月を求めるため、月の日数には依存しない。
    """
    # 12月の各業務日 → 12月 + 翌年1月
    for day in (5, 15, 25, 31):
        assert _target_months(
            datetime.date(2026, 12, day), rolling_window_days=7,
        ) == ["2026-12", "2027-01"]
    # 2月1日（平年・28日）
    assert _target_months(
        datetime.date(2026, 2, 1), rolling_window_days=31,
    ) == ["2026-02", "2026-03"]
    # 2月1日（閏年・29日）
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
    """初回実行: 8 行（当月/来月 × 標準/上位 × 積み上げ/延期）構成で集計される。

    ``[FILTER] PLAN_PREFIXES = [標準, 上位]`` なので、行数は 8 固定。
    ローリングモードの設定は無し（年次累積モード既定）なので、「来月」行は
    構造上存在するが実データは入らない。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    # 2/28 の比較相手として 2025/12/20 を置く（範囲外だが例外的に開かれる）
    make_book(
        input_folder / "一覧_20251220.xlsx",
        [["stay", "2026-04-10", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260228.xlsx",
        [["stay", "2026-04-10", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260420.xlsx",
        [["stay", "2026-04-10", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260422.xlsx",
        [
            ["stay", "2026-04-10", "標準A", "教育", "完了"],
            ["add", "2026-04-12", "標準C", "教育", "予定"],
        ],
    )
    make_book(
        input_folder / "一覧_20260423.xlsx",
        [["stay", "2026-04-10", "標準A", "教育", "完了"]],
    )

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 行は常に 8 行（当月/来月 × 標準/上位 × 積み上げ/延期）
    assert len(rows) == 8
    labels = [_row_label(r) for r in rows]
    assert labels == [
        ("当月", "標準", "教育", "積み上げ"),
        ("当月", "標準", "教育", "延期"),
        ("当月", "上位", "教育", "積み上げ"),
        ("当月", "上位", "教育", "延期"),
        ("来月", "標準", "教育", "積み上げ"),
        ("来月", "標準", "教育", "延期"),
        ("来月", "上位", "教育", "積み上げ"),
        ("来月", "上位", "教育", "延期"),
    ]

    # 業務日 4/21 (一覧_20260421.xlsx が無いので 一覧_20260420.xlsx vs 一覧_20260422.xlsx の比較)
    # → add が 4/22 ファイルに新しく出現 → 業務日 4/21 に積み上げ 1（標準）
    assert _cell(rows, "当月", "積み上げ", "2026-04-21", plan="標準") == "1"
    assert _cell(rows, "当月", "延期", "2026-04-21", plan="標準") == "0"
    # 上位には変化が無い → 該当行は "0"
    assert _cell(rows, "当月", "積み上げ", "2026-04-21", plan="上位") == "0"
    assert _cell(rows, "当月", "延期", "2026-04-21", plan="上位") == "0"
    # 業務日 4/22 (一覧_20260422.xlsx vs 一覧_20260423.xlsx)
    # → add が 4/23 ファイルから消えた → 業務日 4/22 に延期 1（標準）
    assert _cell(rows, "当月", "積み上げ", "2026-04-22", plan="標準") == "0"
    assert _cell(rows, "当月", "延期", "2026-04-22", plan="標準") == "1"


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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    # 業務日 4/22 のためのファイル（=ファイル日付 4/23）はあるが、業務日 4/21 のための
    # ファイル（=ファイル日付 4/22）が無い
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

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
    """増分実行でも行は ``PLAN_PREFIXES`` の件数ぶんの固定行数になる。"""
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    # 1回目: 2025/12/20（2/28 の比較相手）, 2/28, 4/20, 4/21 を実行
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    run()
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    # 業務日 4/20 列 = 一覧_20260421.xlsx vs 一覧_20260420.xlsx の比較結果
    first_4_20 = _cell(first_rows, "当月", "積み上げ", "2026-04-20")

    # 2回目: 4/22, 4/23 のファイルを追加して再実行
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    run()

    with CSV(output_folder / OUTPUT_NAME) as csv:
        second_rows = list(csv.read())
    # 行は常に固定（PLAN_PREFIXES が 2 件のとき 8 行）
    assert len(second_rows) == 8
    # 業務日 4/20 の値が変わらない
    assert _cell(second_rows, "当月", "積み上げ", "2026-04-20") == first_4_20
    # 業務日 4/21, 4/22 が追加される（ファイル日付 4/22, 4/23 でそれぞれ作られる）
    assert "2026-04-21" in second_rows[0]
    assert "2026-04-22" in second_rows[0]
    # 業務日 4/23 は対応するファイル（ファイル日付 4/24）が無いので列が無い
    assert "2026-04-23" not in second_rows[0]
    # 同じキー (当月, 標準, 積み上げ) の行が1つだけ（マージされている）
    matching = [r for r in second_rows
                if r["対象月"] == "当月"
                and r["種別"] == "標準"
                and r["判定"] == "積み上げ"]
    assert len(matching) == 1


def test_run_row_has_values_only_for_target_month_dates(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ラベル ``当月`` 行は、当月だった業務日の列にだけ値を持つ。"""
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 1, 25)
    )
    # 業務日 1/22 まで: 1月の対象月のみ
    make_book(input_folder / "一覧_20251220.xlsx", [["x", "2026-01-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20251231.xlsx", [["x", "2026-01-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260122.xlsx", [["x", "2026-01-10", "標準A", "教育", "完了"]])
    # 業務日 1/22 = ファイル 1/23。x が消える
    make_book(input_folder / "一覧_20260123.xlsx", [])

    # 業務日 1/23 = ファイル 1/24。jan が増える（1月の対象月のみ）
    make_book(
        input_folder / "一覧_20260124.xlsx",
        [
            ["jan", "2026-01-15", "標準A", "教育", "完了"],
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


def _row_label(row: dict[str, object]) -> tuple[str, str, str, str]:
    return (
        str(row["対象月"]),
        str(row["種別"]),
        str(row["作業班"]),
        str(row["判定"]),
    )


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
    make_book(input_folder / "一覧_20251220.xlsx", [["x", "2026-01-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260130.xlsx", [["x", "2026-01-10", "標準A", "教育", "完了"]])
    # 月末のファイル（1/31 終了時点 → 業務日 1/30 で使う）
    make_book(input_folder / "一覧_20260131.xlsx", [["x", "2026-01-10", "標準A", "教育", "完了"]])
    # 月またぎ。一覧_20260201.xlsx は:
    #   - 業務日 1/31 (= 1/31 file vs 2/1 file) の current → 対象月 1月
    #   - 業務日 2/1  (= 2/1 file vs 2/2 file) の previous → 対象月 2月
    make_book(
        input_folder / "一覧_20260201.xlsx",
        [
            ["jan", "2026-01-15", "標準A", "教育", "完了"],
            ["feb", "2026-02-05", "標準B", "予定"],
        ],
    )
    # 2/2 ファイル（業務日 2/1 の current）
    make_book(
        input_folder / "一覧_20260202.xlsx",
        [
            ["jan", "2026-01-15", "標準A", "教育", "完了"],
            ["feb", "2026-02-05", "標準B", "予定"],
        ],
    )

    opened: list[Path] = []

    def _capture(path: Path, *_args: object, **_kwargs: object) -> dict[str, object]:
        opened.append(path)
        return {"jan": __import__("src.source", fromlist=["Record"]).Record(
            "jan", datetime.date(2026, 1, 15), "標準", "教育"),
            "feb": __import__("src.source", fromlist=["Record"]).Record(
                "feb", datetime.date(2026, 2, 5), "標準", "教育"),
            "x": __import__("src.source", fromlist=["Record"]).Record(
                "x", datetime.date(2026, 1, 10), "標準", "教育"),
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
    make_book(
        input_folder / "一覧_20251201.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20251205.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-01-05", "標準A", "教育", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["a", "2026-01-05", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260112.xlsx", [["a", "2026-01-05", "標準A", "教育", "完了"]])

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
    make_book(
        input_folder / "一覧_20251215.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20251218.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20251220.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["a", "2026-01-05", "標準A", "教育", "完了"]])

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
    make_book(
        input_folder / "一覧_20251201.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20251205.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["x", "2026-01-05", "標準A", "教育", "完了"]])
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["x", "2026-01-05", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260112.xlsx", [["x", "2026-01-05", "標準A", "教育", "完了"]])

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
    make_book(
        input_folder / "一覧_20251215.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20251218.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20251220.xlsx",
        [["old", "2026-01-05", "標準A", "教育", "完了"]],
    )
    # 範囲内
    make_book(input_folder / "一覧_20260110.xlsx", [["a", "2026-01-05", "標準A", "教育", "完了"]])

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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-05", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260201.xlsx", [["a", "2026-04-05", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260405.xlsx", [["a", "2026-04-05", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260408.xlsx", [["a", "2026-04-05", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260410.xlsx", [["a", "2026-04-05", "標準A", "教育", "完了"]])
    # 今日より後
    make_book(input_folder / "一覧_20260415.xlsx", [["a", "2026-04-05", "標準A", "教育", "完了"]])

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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2025-12-20", "標準A", "教育", "完了"]])

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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

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
    date_headers = [h for h in headers if h not in ("対象月", "種別", "作業班", "判定")]
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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

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
    first_headers = [
        h for h in first_rows[0].keys()
        if h not in ("対象月", "種別", "作業班", "判定")
    ]
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
    second_headers = [
        h for h in second_rows[0].keys()
        if h not in ("対象月", "種別", "作業班", "判定")
    ]
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
    make_book(
        input_folder / "一覧_20251220.xlsx",
        [["id_a", "2026-08-24", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260824.xlsx",
        [["id_a", "2026-08-24", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260825.xlsx",
        [
            ["id_a", "2026-08-24", "標準A", "教育", "完了"],
            ["id_b", "2026-08-25", "標準B", "教育", "予定"],
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
    make_book(
        input_folder / "一覧_20251220.xlsx",
        [["id_a", "2026-08-23", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260823.xlsx",
        [["id_a", "2026-08-23", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260824.xlsx",
        [["id_a", "2026-08-23", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260825.xlsx",
        [["id_a", "2026-08-23", "標準A", "教育", "完了"]],
    )
    run()
    with CSV(output_folder / OUTPUT_NAME) as csv:
        first_rows = list(csv.read())
    last_date_header = max(
        h for h in first_rows[0].keys() if h not in ("対象月", "種別", "作業班", "判定")
    )
    assert last_date_header == "2026-08-24"

    # 翌日分を追加して再実行
    make_book(
        input_folder / "一覧_20260826.xlsx",
        [["id_a", "2026-08-23", "標準A", "教育", "完了"]],
    )

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
    make_book(
        input_folder / "一覧_20251220.xlsx",
        [["id_a", "2026-08-28", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260828.xlsx",
        [["id_a", "2026-08-28", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260829.xlsx",
        [["id_a", "2026-08-28", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260830.xlsx",
        [["id_a", "2026-08-28", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260831.xlsx",
        [["id_a", "2026-08-28", "標準A", "教育", "完了"]],
    )
    make_book(
        input_folder / "一覧_20260901.xlsx",
        [["id_a", "2026-08-28", "標準A", "教育", "完了"]],
    )

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "種別", "作業班", "判定")]

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
    plan: str = "標準",
    crew: str = "教育",
) -> str:
    row = _row(rows, label, status, plan, crew)
    value = row.get(date_header, "")
    return str(value)


def _row(
    rows: list[dict[str, object]],
    label: str,
    status: str,
    plan: str = "標準",
    crew: str = "教育",
) -> dict[str, object]:
    for row in rows:
        if (
            row["対象月"] == label
            and row["判定"] == status
            and row["種別"] == plan
            and row["作業班"] == crew
        ):
            return row
    raise AssertionError(
        f"行が見つかりません: {label} {plan} {crew} {status}"
    )


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
    make_book(input_folder / "一覧_20260418.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260419.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260421.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "種別", "作業班", "判定")]
    # 窓の下限 = 業務日 4/19。今日が 4/25 で、今日ファイル（= 4/26）は無いので
    # 比較が成立するのは 4/19・4/20・4/21。保存時は start_date から last_compared
    # までの range を必ず列に出す
    assert "2026-04-19" in date_headers
    assert "2026-04-20" in date_headers
    assert "2026-04-21" in date_headers
    # 業務日 4/18 未満の列は作られない（窓の外）
    assert "2026-04-18" not in date_headers
    assert "2026-04-17" not in date_headers
    # 行は ``PLAN_PREFIXES`` の件数ぶんの固定行数（= 8 行）
    assert len(rows) == 8


def test_rolling_mode_year_cumulative_keeps_next_month_row_even_when_empty(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ローリングモードでない年次累積実行では「来月」行も常に存在するが、月の前半は
    lookahead が起きないので実データは入らない（構造上は空のまま）。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 15)
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260413.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260414.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260415.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    labels = [
        (r["対象月"], r["種別"], r["作業班"], r["判定"]) for r in rows
    ]
    # 8 行固定（当月/来月 × {標準, 上位} × 教育 × 積み上げ/延期）
    assert labels == [
        ("当月", "標準", "教育", "積み上げ"),
        ("当月", "標準", "教育", "延期"),
        ("当月", "上位", "教育", "積み上げ"),
        ("当月", "上位", "教育", "延期"),
        ("来月", "標準", "教育", "積み上げ"),
        ("来月", "標準", "教育", "延期"),
        ("来月", "上位", "教育", "積み上げ"),
        ("来月", "上位", "教育", "延期"),
    ]
    # 年次累積モードで月の前半 → 「来月」行は構造上あるが、月初の業務日しか
    # 比較していないので実データなし（対象月はすべて 4月のまま = 「当月」）
    next_added = _row(rows, "来月", "積み上げ", plan="標準")
    next_postponed = _row(rows, "来月", "延期", plan="標準")
    # データが無い業務日列は空セル、比較対象で対象月だったら "0"。
    # 月の途中で全て 4 月以外の業務日がなければ「来月」の対象月は発生しない
    # → 対象月の列は "0" / "" が混在で、件数セルとして 0 や "" が入る
    # 「来月」行（積み上げ・延期の双方）が **4 月の業務日に件数を持っていない**ことを確認
    for row_dict in (next_added, next_postponed):
        for header, value in row_dict.items():
            if header in ("対象月", "種別", "作業班", "判定"):
                continue
            assert value == "0" or value == ""


def test_rolling_mode_writes_next_month_row_for_next_month_cases(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ローリングモードでは、月のどの業務日でも翌月の案件が「来月」行に反映される。

    今日=4/25・窓 7。ローリングモードでは無条件で当月 + 翌月の2要素を返すので、
    4/25 の業務日の比較結果（4/25 vs 一覧_20260426.xlsx = 翌日ファイル）は、
    4月案件が「当月（4月）」行に、5月案件が「来月（5月）」行にそれぞれ入る。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25), rolling_window_days=7,
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260419.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(
        input_folder / "一覧_20260425.xlsx",
        [
            ["a", "2026-04-10", "標準A", "教育", "完了"],
            ["new_apr", "2026-04-15", "標準A", "教育", "完了"],
            ["new_may", "2026-05-02", "標準A", "教育", "完了"],
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


def test_rolling_mode_writes_next_month_row_early_in_month(
    tmp_path: Path, make_book, setup_run
) -> None:
    """ローリングモードでは、月初の業務日でも翌月の案件が「来月」行に入る。

    旧仕様では「月末 N 日前」だけ lookahead していたため、4/10 では「来月」行が
    常に空だった。新仕様では無条件で当月 + 翌月を返すため、月の前半でも
    翌月（5月）の案件が増えていれば「来月」行に反映される。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 10), rolling_window_days=7,
    )
    # 4/3 の比較相手として 2025/12/20 を置く（範囲外だが例外的に読まれる）
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260403.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    # 業務日 4/9（ファイル日付 4/10）で、翌月の案件 new_may が増える
    make_book(
        input_folder / "一覧_20260410.xlsx",
        [
            ["a", "2026-04-10", "標準A", "教育", "完了"],
            ["new_apr", "2026-04-15", "標準A", "教育", "完了"],
            ["new_may", "2026-05-02", "標準A", "教育", "完了"],
        ],
    )

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 業務日 4/9 列 = 一覧_20260410.xlsx vs 一覧_20260403.xlsx の比較結果
    # 当月の行: 4月案件 new_apr が増えた → 当月・積み上げ = 1
    # 来月の行: 5月案件 new_may が増えた → 来月・積み上げ = 1
    assert _cell(rows, "当月", "積み上げ", "2026-04-09") == "1"
    assert _cell(rows, "来月", "積み上げ", "2026-04-09") == "1"
    assert _cell(rows, "当月", "延期", "2026-04-09") == "0"
    assert _cell(rows, "来月", "延期", "2026-04-09") == "0"


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
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260102.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())
    headers = list(rows[0].keys())
    date_headers = [h for h in headers if h not in ("対象月", "種別", "作業班", "判定")]
    # 年次累積モード: 今年の 1/1 〜 今日（4/25）が対象
    assert any(h.startswith("2026-01") for h in date_headers), (
        f"1月の列が存在すること（年次累積）: {date_headers}"
    )
    # 4月の列も出ている
    assert any(h.startswith("2026-04") for h in date_headers), (
        f"4月の列が存在すること（年次累積）: {date_headers}"
    )
    # 行は ``PLAN_PREFIXES`` の件数ぶんの固定行数（= 8 行）
    assert len(rows) == 8


def test_run_writes_eight_rows_for_two_plan_prefixes(
    tmp_path: Path, make_book, setup_run
) -> None:
    """``[FILTER] PLAN_PREFIXES = [標準, 上位]`` のとき、8 行
    （= 当月/来月 × 2 種別 × 2 判定）が固定で並ぶ。

    合算ではなく種別ごとに別行として並ぶため、``来月, 上位, 積み上げ`` などの
    4 種以外の行も CSV に存在する。
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [
        ["a", "2026-04-10", "標準A", "教育", "完了"],
        ["b", "2026-04-15", "上位B", "教育", "完了"],
    ])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 8 行固定（当月/来月 × 標準/上位 × 教育 × 積み上げ/延期）
    assert len(rows) == 8
    labels = [_row_label(r) for r in rows]
    assert labels == [
        ("当月", "標準", "教育", "積み上げ"),
        ("当月", "標準", "教育", "延期"),
        ("当月", "上位", "教育", "積み上げ"),
        ("当月", "上位", "教育", "延期"),
        ("来月", "標準", "教育", "積み上げ"),
        ("来月", "標準", "教育", "延期"),
        ("来月", "上位", "教育", "積み上げ"),
        ("来月", "上位", "教育", "延期"),
    ]
    # 「当月・標準・積み上げ」と「当月・上位・積み上げ」が別のセルとして存在すること
    # （= 種別をまたいで合算されていない）
    std = _row(rows, "当月", "積み上げ", plan="標準")
    hi = _row(rows, "当月", "積み上げ", plan="上位")
    assert std is not hi
    # 4/22 (一覧_20260423.xlsx - 一覧_20260420.xlsx, 直前の 4/20 ファイル比較) →
    # "b" が新規 → 当月・上位・積み上げ = 1、当月・標準・積み上げ = 0
    assert _cell(rows, "当月", "積み上げ", "2026-04-22", plan="標準") == "0"
    assert _cell(rows, "当月", "積み上げ", "2026-04-22", plan="上位") == "1"


def test_run_row_keys_depend_on_plan_prefixes_length(
    tmp_path: Path, make_book, setup_run
) -> None:
    """``PLAN_PREFIXES`` の件数に応じて行数が変わる（= 「常に8行」ではない）。

    1件に絞り込めば 4 行（= 旧来の構成）に戻ることを確認する。これは
    「``PLAN_PREFIXES = [標準, 上位]`` のとき 8 行」とセットで確認することで、
    「種別数 × 2 × 2」の構造になっていることを担保する。
    """
    input_folder, output_folder = setup_run(
        tmp_path,
        today_date=datetime.date(2026, 4, 25),
        kinds_value="[完了]",
        plan_prefixes_value="[標準]",
    )
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 1 種別 × 1 作業班 × 2 対象月 × 2 判定 = 4 行
    assert len(rows) == 4
    labels = [_row_label(r) for r in rows]
    assert labels == [
        ("当月", "標準", "教育", "積み上げ"),
        ("当月", "標準", "教育", "延期"),
        ("来月", "標準", "教育", "積み上げ"),
        ("来月", "標準", "教育", "延期"),
    ]


def test_run_row_keys_depend_on_crews_length(
    tmp_path: Path, make_book, setup_run
) -> None:
    """``CREWS`` の件数に応じて行数が変わる（= 「常に8行」ではない）。

    ``[FILTER] CREWS = [教育]`` のとき 8 行 = 「``PLAN_PREFIXES = [標準, 上位]`` の
    とき 8 行」とセットで確認することで、「種別数 × 作業班数 × 2 × 2」の構造に
    なっていることを担保する。
    """
    input_folder, output_folder = setup_run(
        tmp_path,
        today_date=datetime.date(2026, 4, 25),
        crews_value="[教育, 作業班B]",
    )
    # 作業班2種 × 種別1種 × 2 対象月 × 2 判定 = 8 行 を検証するため、
    # 種別は1件に絞る（``plan_prefixes_value="[標準]"``）
    # 注: setup_run の plan_prefixes_value を絞ったケースで再セットアップするため、
    # もう一度 setup_run を呼び出して。出力しないことがないように
    pass  # 上記の crews_value 指定だけでは plan_prefixes_value がデフォルトのままなので、
    # 作業班2種 × 種別2種 × 2 対象月 × 2 判定 = 16 行 を確認する形に切り替える
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260423.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

    run()

    csv_path = output_folder / OUTPUT_NAME
    with CSV(csv_path) as csv:
        rows = list(csv.read())

    # 2 種別 × 2 作業班 × 2 対象月 × 2 判定 = 16 行
    assert len(rows) == 16
    # 4 種のキー列がすべて並んでいること（「当月, 標準, 教育, 積み上げ」など）
    seen_labels = {_row_label(r) for r in rows}
    assert ("当月", "標準", "教育", "積み上げ") in seen_labels
    assert ("当月", "上位", "教育", "積み上げ") in seen_labels
    assert ("当月", "標準", "作業班B", "積み上げ") in seen_labels
    assert ("当月", "上位", "作業班B", "積み上げ") in seen_labels
    # 同じキーが重複しない（8 行構成で複数キーが同じ行を共有しない）
    assert len(seen_labels) == 16


def test_run_logs_newest_file_date_with_corresponding_business_date(
    tmp_path: Path, make_book, setup_run, caplog
) -> None:
    """読み込んだファイル群のうち最新のファイル日付と、そのファイルが表す業務日をログに出す。

    集計表の横軸は **業務日**（= ファイル日付 - 1日）なので、今日=4/25 実行のとき
    最新のファイルは ``一覧_20260424.xlsx`` で、業務日 4/23 終了時点のデータになる。
    この事実を実行ログから直接読み取れるよう、新しいログ行を追加した。
    ファイル日付の値と、それが表す業務日の値が両方ログに含まれていることを assert する。
    """
    input_folder, _output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 4, 25)
    )
    # 4/20 ファイルの比較相手として 2025/12/20 を置く
    make_book(input_folder / "一覧_20251220.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260228.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260420.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    make_book(input_folder / "一覧_20260422.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])
    # 今日=4/25 の中で最新のファイル。業務日 4/23 終了時点のデータ
    make_book(input_folder / "一覧_20260424.xlsx", [["a", "2026-04-10", "標準A", "教育", "完了"]])

    with caplog.at_level(logging.INFO):
        run()

    # ログに「読み込んだファイルの最新日付」という文言と、ファイル日付 4/24、
    # そのファイルが表す業務日 4/23 が両方含まれている
    assert "読み込んだファイルの最新日付" in caplog.text
    assert "2026-04-24" in caplog.text
    assert "2026-04-23" in caplog.text


def test_run_logs_when_already_up_to_date_and_skips_calculation(
    tmp_path: Path, make_book, setup_run, caplog
) -> None:
    """同日次に計算すべき業務日 ``start_date`` が、実際に計算可能な上限 ``end_date``
    を超えているとき（= 「次のファイルが来るのを待っているだけ」の日常的な状態）、
    専用の info ログを出して何も計算せずに終わる。

    検証:
    - 「既に最新の業務日」という文言がログに出る
    - last_date と「次に必要になるファイル日付」の両方がログに出る
    - 出力 CSV は実行前後で変化しない（新規計算が起きない）
    """
    input_folder, output_folder = setup_run(
        tmp_path, today_date=datetime.date(2026, 8, 27)
    )

    # 既存 ``集計.csv`` を最終業務日 2026-08-26 の状態で作る。
    # ``write_csv`` を直接呼ぶことで、本物の「増分済み」状態を再現する
    # （テスト用の手書き CSV だと列見出しが parse できない／日付が反映されない
    # 恐れがあるため、``write_csv`` 経由で書く）。
    from src.report import write_csv  # テストローカル import に統一
    plan_prefixes = ("標準", "上位")
    crews = ("教育",)
    row_keys = [
        (label, plan, crew)
        for label in ROW_LABELS
        for plan in plan_prefixes
        for crew in crews
    ]
    last_business_date = datetime.date(2026, 8, 26)
    csv_path = output_folder / OUTPUT_NAME
    write_csv(
        csv_path,
        by_row={},
        dates=[(last_business_date, True)],
        row_keys=row_keys,
    )

    # 実行前の CSV を丸ごと覚えておく
    with CSV(csv_path) as csv:
        before_rows = list(csv.read())
    headers_before = list(before_rows[0].keys())

    # 入力フォルダには今日のファイル ``一覧_20260827.xlsx`` だけを置く。
    # 明日のファイル（一覧_20260828.xlsx）はまだ無い想定 → 業務日 8/27 を
    # 計算するためのファイル日付 8/28 が無い状態。
    make_book(input_folder / "一覧_20260827.xlsx", [["a", "2026-08-01", "標準A", "教育", "完了"]])

    with caplog.at_level(logging.INFO):
        run()

    # 1. 「既に X 終了時点まで計算済みです」という文言が出ている
    assert "既に 2026-08-26 終了時点まで計算済み" in caplog.text
    # 2. last_date（2026-08-26）と、次に必要なファイル日付（2026-08-28）が両方ログにある
    assert "2026-08-26" in caplog.text
    assert "2026-08-28" in caplog.text
    # 3. CSV の内容が実行前後で変わっていない（列・行の構造と、列見出しの最終業務日が同じ）
    with CSV(csv_path) as csv:
        after_rows = list(csv.read())
    assert [list(r.keys()) for r in after_rows] == [list(r.keys()) for r in before_rows]
    date_headers_after = [
        h for h in after_rows[0].keys() if h not in ("対象月", "種別", "作業班", "判定")
    ]
    date_headers_before = [
        h for h in headers_before if h not in ("対象月", "種別", "作業班", "判定")
    ]
    assert date_headers_after == date_headers_before
    # 念のため：最終業務日列は実行前と同じ 2026-08-26 のまま
    assert date_headers_after[-1] == "2026-08-26"
    # 「新規計算」が起きていないこと（= 2026-08-27 の列は作られていない）
    assert "2026-08-27" not in date_headers_after
