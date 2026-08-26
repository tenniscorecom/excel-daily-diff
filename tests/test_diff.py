import datetime
from pathlib import Path
from unittest.mock import patch

from src.diff import (
    LABEL_CURRENT_MONTH,
    LABEL_NEXT_MONTH,
    STATUS_ADDED,
    STATUS_POSTPONED,
    RowKey,
    compute_counts,
)
from src.source import Record


def _record(customer_id: str, day: int = 1, plan_prefix: str = "標準") -> Record:
    return Record(customer_id, datetime.date(2026, 4, day), plan_prefix)


def _target_months_for(business_date: datetime.date) -> list[str]:
    """テスト用の対象月決定。業務日 = その月の単一対象月（文字列）。"""
    return [f"{business_date.year:04d}-{business_date.month:02d}"]


def test_compute_counts_compares_adjacent_files_and_reads_each_once() -> None:
    paths = [Path(f"一覧_2026042{day}.xlsx") for day in range(3)]
    records = [
        {"a": _record("a", 1), "b": _record("b", 2)},
        {"b": _record("b", 3), "c": _record("c", 4)},
        {"c": _record("c", 5), "d": _record("d", 6)},
    ]
    dated_files = [(datetime.date(2026, 4, 20 + index), path) for index, path in enumerate(paths)]

    with patch("src.diff.read_records", side_effect=records) as reader:
        counts = compute_counts(dated_files, _target_months_for, ("標準",), ("完了",), ())

    added = counts.by_row[RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_ADDED)]
    postponed = counts.by_row[RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_POSTPONED)]
    # 業務日 = ファイル日付 - 1日。
    # 一覧_20260421.xlsx(業務日 4/20) vs 一覧_20260420.xlsx(業務日 4/19):
    #   c は 4/20 にのみ存在 → 業務日 4/20 に積み上げ 1
    #   a は 4/19 にのみ存在 → 業務日 4/20 に延期 1
    assert added[datetime.date(2026, 4, 20)] == 1  # "c"
    assert postponed[datetime.date(2026, 4, 20)] == 1  # "a"
    # 一覧_20260422.xlsx(業務日 4/21) vs 一覧_20260421.xlsx(業務日 4/20):
    #   d は 4/21 にのみ存在 → 業務日 4/21 に積み上げ 1
    #   b は 4/20 にのみ存在 → 業務日 4/21 に延期 1
    assert added[datetime.date(2026, 4, 21)] == 1  # "d"
    assert postponed[datetime.date(2026, 4, 21)] == 1  # "b"
    assert counts.compared_dates == {datetime.date(2026, 4, 20), datetime.date(2026, 4, 21)}
    assert [call.args[0] for call in reader.call_args_list] == paths


def test_compute_counts_only_buckets_into_target_months() -> None:
    """業務日の対象月だけがバケットされる。4/30 → 4月のみ、5/1 → 5月のみ。"""
    paths = [Path(f"一覧_20260{day}.xlsx") for day in (430, 501, 502)]
    # 4/30: 4月のレコード、5/1: 4月と5月のレコードが混在、5/2: 5月のレコード
    records = [
        {
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
        },
        {
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
        {
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),
        (datetime.date(2026, 5, 1), paths[1]),
        (datetime.date(2026, 5, 2), paths[2]),
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("標準",), ("完了",), ()
        )

    # 業務日 4/29 (4/30 ファイル vs previous=無) → 比較なし
    # 業務日 4/30 (5/1 ファイル vs 4/30 ファイル) → 対象月は 4月のみ
    #   4月: apr が両方 → 差分なし
    #   5月: 4/30 ファイル側に 5月のレコードが無いので、5月のバケットは無い
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_ADDED) not in counts.by_row
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_POSTPONED) not in counts.by_row
    # 業務日 5/1 (5/2 ファイル vs 5/1 ファイル) → 対象月は 5月のみ
    #   5月: may が両方 → 差分なし
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_ADDED) not in counts.by_row
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_POSTPONED) not in counts.by_row


def test_compute_counts_separates_rows_by_plan_prefix() -> None:
    """種別（標準/上位）のレコードを別々の行に分けて集計する。

    ``(当月, 標準, 判定)`` と ``(当月, 上位, 判定)`` がそれぞれ別セルとして
    存在し、標準・上位の合算はされない。``PLAN_PREFIXES`` を変えてもこの
    区別は変わらない（``run.py`` 側でこの区別をまとめて消していない）。
    """
    paths = [Path("一覧_20260421.xlsx"), Path("一覧_20260422.xlsx")]
    records = [
        {
            # 4月の標準レコード。両方にある → 差分なし
            "a_std": Record("a_std", datetime.date(2026, 4, 10), "標準"),
            # 4月の上位レコード。4/22 ファイルから消える → 延期 1
            "b_hi": Record("b_hi", datetime.date(2026, 4, 10), "上位"),
        },
        {
            "a_std": Record("a_std", datetime.date(2026, 4, 10), "標準"),
            # 4月の上位レコード。4/22 で新規 → 積み上げ 1
            "c_hi": Record("c_hi", datetime.date(2026, 4, 10), "上位"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 21), paths[0]),
        (datetime.date(2026, 4, 22), paths[1]),
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("標準", "上位"), ("完了",), ()
        )

    # 一覧_20260421.xlsx(業務日 4/20) vs 一覧_20260422.xlsx(業務日 4/21):
    #   標準: 差し引き 0（a_std は両方にある）→ キーが作られない（差分が無いので）
    #   上位: c_hi 追加 +1、b_hi 消失 -1 → (当月, 上位, 積み上げ) = 1、
    #         (当月, 上位, 延期) = 1
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_ADDED) not in counts.by_row
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_POSTPONED) not in counts.by_row
    hi_added = counts.by_row[RowKey(LABEL_CURRENT_MONTH, "上位", STATUS_ADDED)]
    hi_postponed = counts.by_row[RowKey(LABEL_CURRENT_MONTH, "上位", STATUS_POSTPONED)]
    assert hi_added[datetime.date(2026, 4, 21)] == 1
    assert hi_postponed[datetime.date(2026, 4, 21)] == 1
    # キーは全て3要素で、種別ごとに別セルとして存在する
    assert all(len(key) == 3 for key in counts.by_row)
    # 差分の出た 2 セルだけ by_row に登録される（差分が無い行はキー自体が無い）
    assert len(counts.by_row) == 2


def test_compute_counts_row_keys_follow_plan_prefixes_order() -> None:
    """``PLAN_PREFIXES`` の並び順がそのまま by_row のキー順序になる。

    種別は ``compute_counts`` 側で並べ替えず、``read_records`` 側で
    ``PLAN_PREFIXES`` の順にマッチした先頭一致の語が入る（``Record.plan_prefix``）。
    ここでは ``PLAN_PREFIXES = ("上位", "標準")`` の逆順でも、呼び出し側が知る
    順序はそのまま by_row のキー集合に反映されることを確認する。
    """
    paths = [Path("一覧_20260421.xlsx"), Path("一覧_20260422.xlsx")]
    records = [
        {"a": Record("a", datetime.date(2026, 4, 10), "上位")},
        {"b": Record("b", datetime.date(2026, 4, 10), "上位")},
    ]
    dated_files = [
        (datetime.date(2026, 4, 21), paths[0]),
        (datetime.date(2026, 4, 22), paths[1]),
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("上位", "標準"), ("完了",), ()
        )

    # 上位の行（= (当月, 上位, 積み上げ) ）に 1 件
    assert counts.by_row[RowKey(LABEL_CURRENT_MONTH, "上位", STATUS_ADDED)][
        datetime.date(2026, 4, 21)
    ] == 1
    # 標準はレコードが無いためキーが作られない
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_ADDED) not in counts.by_row


def test_compute_counts_uses_business_date_target_months_across_month_boundary() -> None:
    """月をまたぐ業務日（4/30 → 5/1）で、各業務日の対象月が変わる。

    業務日 4/29 → 4月、業務日 4/30 → 4月、業務日 5/1 → 5月。
    一覧_20260501.xlsx は 4月対象と 5月対象の両方でバケットされる。
    """
    paths = [
        Path("一覧_20260430.xlsx"),
        Path("一覧_20260501.xlsx"),
        Path("一覧_20260502.xlsx"),
    ]
    records = [
        {
            # 4月のレコード
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
        },
        {
            # 4月のレコード（apr は引き続き残存）
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
        },
        {
            # 5月のレコード（may は 5月 bucket で新規に登場）
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),  # → 業務日 4/29
        (datetime.date(2026, 5, 1), paths[1]),   # → 業務日 4/30
        (datetime.date(2026, 5, 2), paths[2]),   # → 業務日 5/1
    ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_for, ("標準",), ("完了",), ()
        )

    # 業務日 4/29 (4/30 ファイル vs previous=無し) → 比較なし
    # 業務日 4/30 (5/1 ファイル vs 4/30 ファイル) → 対象月は 4月のみ（apr が両側 → 差分なし）
    # 業務日 5/1 (5/2 ファイル vs 5/1 ファイル) → 対象月は 5月のみ
    #   previous_by_month["2026-05"] は 5/1 ファイル側の 5月 bucket（空）。
    #   5/2 ファイルの "2026-05" bucket に may がある → 積み上げ 1
    # 延期は一切発生しない
    assert RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_POSTPONED) not in counts.by_row
    # 積み上げキーは業務日 5/1 で 1
    assert counts.by_row[RowKey(LABEL_CURRENT_MONTH, "標準", STATUS_ADDED)][
        datetime.date(2026, 5, 1)
    ] == 1


def test_compute_counts_reads_each_file_only_once_across_month_boundary() -> None:
    """月をまたぐ日の同じファイルが、previous/current の両方で1回しか読まれない。"""
    paths = [
        Path("一覧_20260430.xlsx"),
        Path("一覧_20260501.xlsx"),
        Path("一覧_20260502.xlsx"),
    ]
    records = [
        {"a": Record("a", datetime.date(2026, 4, 30), "標準")},
        {"a": Record("a", datetime.date(2026, 4, 30), "標準"), "b": Record("b", datetime.date(2026, 5, 1), "標準")},
        {"a": Record("a", datetime.date(2026, 4, 30), "標準"), "b": Record("b", datetime.date(2026, 5, 1), "標準")},
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),  # → 業務日 4/29, 4/30
        (datetime.date(2026, 5, 1), paths[1]),   # → 業務日 4/30, 5/1
        (datetime.date(2026, 5, 2), paths[2]),   # → 業務日 5/1
    ]

    with patch("src.diff.read_records", side_effect=records) as reader:
        compute_counts(dated_files, _target_months_for, ("標準",), ("完了",), ())

    # 各ファイルが読まれるのは1回だけ
    opened = [call.args[0] for call in reader.call_args_list]
    assert opened == paths
    assert len(opened) == len(set(opened))


def test_compute_counts_emits_next_month_label_for_lookahead() -> None:
    """``target_months_for`` が「当月 + 翌月」を返すとき、``来月`` ラベル行が現れる。

    ローリングモードの月末 lookahead に相当するケース。
    1番目（=業務日自身の月）が ``当月``、2番目（=翌月）が ``来月`` に変換される。
    """
    paths = [Path("一覧_20260430.xlsx"), Path("一覧_20260501.xlsx")]
    records = [
        {
            # 4月のレコード
            "apr": Record("apr", datetime.date(2026, 4, 30), "標準"),
        },
        {
            # 5月のレコード（先月ファイルには無い）
            "may": Record("may", datetime.date(2026, 5, 10), "標準"),
        },
    ]
    dated_files = [
        (datetime.date(2026, 4, 30), paths[0]),  # → 業務日 4/29
        (datetime.date(2026, 5, 1), paths[1]),   # → 業務日 4/30
    ]

    def _target_months_with_next(business_date: datetime.date) -> list[str]:
        return [
            f"{business_date.year:04d}-{business_date.month:02d}",
            f"{business_date.year if business_date.month != 12 else business_date.year + 1}"
            f"-{(business_date.month % 12) + 1:02d}",
        ]

    with patch("src.diff.read_records", side_effect=records):
        counts = compute_counts(
            dated_files, _target_months_with_next, ("標準",), ("完了",), ()
        )

    # 業務日 4/30 で「来月 = 5月」も対象に入る。
    # 4/30 ファイルには 5月のレコードが無いが、5/1 ファイルには may がある
    # → (来月, 標準, 積み上げ) = 1
    assert counts.by_row[RowKey(LABEL_NEXT_MONTH, "標準", STATUS_ADDED)][
        datetime.date(2026, 4, 30)
    ] == 1
    # 「来月」が対象だった業務日として記録される
    assert datetime.date(2026, 4, 30) in counts.target_dates_by_month[LABEL_NEXT_MONTH]
    # 「当月」も同様に記録される
    assert datetime.date(2026, 4, 30) in counts.target_dates_by_month[LABEL_CURRENT_MONTH]
