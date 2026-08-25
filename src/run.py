"""
src/run.py — 処理の本体

実行日→既存 CSV→期間の始まり、入力フォルダ→期間の終わり を組み合わせて
集計範囲を決め、日ごとの延期・積み上げを集計する。

日付が2種類出てくるので混同しないこと。
    ファイル名の日付 … いつ時点の一覧かを表す（``一覧_YYYYMMDD.xlsx`` の YYYYMMDD）
    案件の日付      … 絞り込みにだけ使う。対象月に入っているかを見る

集計表の横軸は **業務日**（= ファイル名の日付 - 1日）。入力ファイルは
「前日終了時点」のデータなので、``一覧_20260825.xlsx`` の中身は 8/24 終了時点の
状態であり、``一覧_20260824.xlsx`` との差分は「8/24 に動いたぶん」になる。

**対象月は業務日基準**。業務日 D の列は、D の月で案件を絞る。実行日が
8 月でも、業務日 1 月の列は 1 月の案件だけを対象にする（= 1 月時点の
積み上げ・延期を 1 月の案件について見る、という集計の意図に沿う）。
"""

import datetime
import logging
from pathlib import Path

from comken import config
from comken.core import DateFileFinder, date_in_name, today

from src.diff import ByRow, DaySavedCallback, compute_counts
from src.report import COL_TARGET_MONTH, last_date_in_csv, read_existing, write_csv
from src.source import load_rules

logger = logging.getLogger(__name__)

# 出力ファイル名（[REPORT] OUTPUT_FOLDER の下に固定名で置く。日付入りにはしない）
OUTPUT_NAME = "集計.csv"


def run() -> Path:
    """集計を実行し、出力先 CSV のパスを返す。"""
    today_date = today()

    plan_prefixes = tuple(config.FILTER.PLAN_PREFIXES)
    kinds = tuple(config.FILTER.KINDS)
    rules = load_rules()
    input_folder = Path(config.FILES.INPUT_FOLDER)
    output_folder = Path(config.REPORT.OUTPUT_FOLDER)
    output_path = output_folder / OUTPUT_NAME

    dated_files = _all_dated_files(input_folder)
    if not dated_files:
        logger.warning("入力フォルダに対象ファイルがありません: %s", input_folder)
        return output_path

    existing = read_existing(output_path)

    # 下限：実行日の年の1月1日（業務日）。それより古いファイルは対象外
    # （古いファイルはシート構造が違うことがある）
    range_floor = _range_floor(today_date)
    # 上限：実行日（業務日）。今日より後の日付のファイルは対象外
    range_ceiling = today_date

    last_date = last_date_in_csv(existing) if existing else None
    if last_date is None:
        # 全期間を作り直す
        start_date = range_floor
        run_mode = "初回"
    else:
        # 増分計算。CSV の最後の業務日が D のとき、次に追加する列は業務日 D+1。
        # 業務日 D+1 の列を作るには「ファイル日付 D+2」のファイル（D+1 終了時点）が要る。
        # 業務日基準の ``start_date`` は D+1（= ``last_date + 1日``）。``_files_in_range``
        # の中でファイル日付 D+2 起点に変換される。
        start_date = max(last_date + datetime.timedelta(days=1), range_floor)
        run_mode = "増分"

    # 業務日の end_date = dated_files の最新ファイル日付 - 1日
    end_date = min(
        dated_files[-1][0] - datetime.timedelta(days=1),
        range_ceiling,
    )
    targets, has_predecessor = _files_in_range(dated_files, start_date, end_date)
    if not targets:
        logger.warning(
            "読み込み対象の範囲（%s 〜 %s）にファイルがありません: %s",
            start_date,
            end_date,
            input_folder,
        )
        return output_path
    if not has_predecessor:
        logger.warning(
            "範囲内ファイルの比較相手が範囲外にも存在しないため、%s ぶんは集計できません",
            start_date,
        )
        return output_path

    # 業務日基準の対象月。読み込み範囲全体で登場する対象月の和集合を取る。
    # 実行日基準の単一グループではないので、開始ログは「範囲内で出現しうる対象月」
    # の全体を出す（業務日ごとのログにはその日に使った対象月が出る）。
    target_months_in_range = _target_months_in_range(start_date, end_date)

    # 実行開始時の概要ログ。処理の進行に合わせて直書きする（途中でどこまで
    # 進んでいたかがログから追えるように）。
    logger.info("入力フォルダ: %s（パターン: %s）", input_folder, config.FILES.FILE_PATTERN)
    logger.info("対象月: %s", ", ".join(target_months_in_range))
    logger.info("実行モード: %s", run_mode)
    logger.info("読み込み範囲（業務日）: %s 〜 %s", start_date, end_date)
    # 業務日 start_date に対応するファイル日付は start_date + 1日
    in_range_file_start = start_date + datetime.timedelta(days=1)
    in_range_count = sum(1 for date, _ in targets if date >= in_range_file_start)
    skipped_count = len(dated_files) - in_range_count
    summary = (
        f"入力フォルダのファイル {len(dated_files)} 件のうち、"
        f"{in_range_count} 件を読み込み、{skipped_count} 件は範囲外で読み飛ばし"
    )
    if has_predecessor:
        predecessor_name = targets[0][1].name
        summary += (
            f"。{start_date}（業務日）の比較相手として、"
            f"範囲外から 1 ファイル追加で読みます（{predecessor_name}）"
        )
    else:
        summary += "（比較相手はありません）"
    logger.info("%s", summary)

    # 既存 CSV に残っている日付列（再開時に過去側の列が消えないように）。
    # 毎回保存時に ``start_date`` から「直近で比較が終わった日」までの range と
    # 合併して列見出しを作る（未比較の区間があると再開位置がズレるため）。
    existing_dates = _dates_from_existing(existing)

    # 書き出す行の (対象月, 種別) の組を呼び出し側で組み立てて渡す。
    # 読み込み範囲全体で対象月になる月の行を必ず作り、既存 CSV にある対象月も
    # 残す（古い対象月の行を消さない）。
    existing_months = _months_from_existing(existing)
    months_for_rows = sorted(set(target_months_in_range) | existing_months)
    row_keys = [(month, plan) for month in months_for_rows for plan in plan_prefixes]

    save_callback, save_sizes = _make_day_saver(
        output_path=output_path,
        start_date=start_date,
        existing_dates=existing_dates,
        row_keys=row_keys,
    )
    # 1日ぶん終わるたびに CSV を途中保存することを最初に1回だけ伝える。
    # 長い処理で「今どこまで保存されているのか」がログから追えるようにするための
    # 既存の方針と整合させる。途中で落ちてもここまでがファイルに残っている。
    logger.info("1日ぶんごとに保存しながら進めます")

    counts = compute_counts(
        targets,
        _target_months,
        plan_prefixes,
        kinds,
        rules,
        range_start=start_date,
        on_day_done=save_callback,
    )

    logger.info(
        "%s 〜 %s のうち %d 日ぶんを比較しました（対象月: %s）",
        start_date,
        end_date,
        len(counts.compared_dates),
        ", ".join(target_months_in_range),
    )
    # 出力ログは最後の1回だけにする。途中保存のたびに毎回出していたものを、
    # 完了後のこの1回に置き換える。最終的な行数と列数は最後の保存の形。
    if save_sizes:
        rows, cols = save_sizes[-1]
        logger.info(
            "出力しました: %s（%d 行 × %d 列）",
            output_path,
            rows,
            cols,
        )
    return output_path


def _range_floor(today_date: datetime.date) -> datetime.date:
    """読み込み範囲の下限。実行日の年の1月1日（業務日）。

    1月実行時でも同年の1月1日（前年の12月1日にはしない）。
    古いファイルはシート構造が違うことが多いため、範囲外との比較相手に
    使う1ファイルを除いて、今年ぶんだけを読む。
    """
    return datetime.date(today_date.year, 1, 1)


def _target_months(business_date: datetime.date) -> list[str]:
    """業務日から対象月を決める。業務日 = その業務日の属する月の ``"YYYY-MM"`` を返す。"""
    return [f"{business_date.year:04d}-{business_date.month:02d}"]


def _target_months_in_range(
    start: datetime.date, end: datetime.date
) -> list[str]:
    """``start`` から ``end`` までの各業務日について対象月を求め、和集合を返す。

    業務日ごとに ``_target_months`` を呼ぶので、月をまたぐ境界があれば
    自動的に複数の対象月が並ぶ。順序は年月昇順。
    """
    months: set[str] = set()
    current = start
    while current <= end:
        months.update(_target_months(current))
        current += datetime.timedelta(days=1)
    return sorted(months)


def _date_range(start: datetime.date, end: datetime.date) -> list[datetime.date]:
    """``start`` から ``end`` までの全日付を古い順で返す（1日も飛ばさない）。"""
    return [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]


def _dates_from_existing(existing: list[dict[str, object]]) -> list[datetime.date]:
    """既存 CSV の1行目から、日付ヘッダとして解釈できたものを抽出する。"""
    if not existing:
        return []
    first_col_index = 3  # 対象月 / 種別 / 判定 の3列のあと
    dates: list[datetime.date] = []
    for header in list(existing[0].keys())[first_col_index:]:
        text = header.strip()
        try:
            dates.append(datetime.date.fromisoformat(text))
        except ValueError:
            continue
    return dates


def _months_from_existing(existing: list[dict[str, object]]) -> set[str]:
    """既存 CSV の「対象月」列を読み取り、``"YYYY-MM"`` の集合を返す。

    既存の対象月の行を残すために使う。解釈できない行はスキップする
    （手編集などで壊れた行があっても落とさず、後段の ``write_csv`` が
    ``leftover_rows`` としてそのまま残すので問題ない）。
    """
    months: set[str] = set()
    for row in existing:
        value = row.get(COL_TARGET_MONTH, "")
        if not value:
            continue
        text = str(value).strip()
        # CSV の対象月は既に ``YYYY-MM`` 形式。空文字でなければそのまま採用
        if text:
            months.add(text)
    return months


def _all_dated_files(folder: Path) -> list[tuple[datetime.date, Path]]:
    """フォルダ内の対象ファイルを、日付の古い順に並べて返す。

    config.ini の ``FILE_PATTERN`` は ``一覧_*.xlsx`` のように ``*`` を含めて
    書く。``DateFileFinder.dated`` は ``*`` をワイルドカード扱いせず文字どおりの
    前方一致で探すので、ここで ``*`` を落として渡す。
    """
    prefix = str(config.FILES.FILE_PATTERN).replace("*", "")
    found = DateFileFinder(folder).dated(prefix)
    return sorted(
        (date, path) for path in found if (date := date_in_name(path.name)) is not None
    )


def _files_in_range(
    dated_files: list[tuple[datetime.date, Path]],
    start: datetime.date,
    end: datetime.date,
) -> tuple[list[tuple[datetime.date, Path]], bool]:
    """期間内 + 範囲の最初のファイルの直前1ファイルを取り出して返す。

    ``start`` と ``end`` は **業務日**。``dated_files`` の ``date`` は **ファイル日付**
    （= 業務日 + 1日）なので、内部で ``+1日`` してファイル日付に変換して比較する。
    戻り値は ``(対象ファイルのリスト, 直前ファイルが範囲外にあったか)``。
    直前ファイルが ``dated_files`` の先頭に達して範囲外に無いときは False を返し、
    そのときは呼び出し側で集計をスキップする（比較相手がいないため）。
    """
    # 業務日 → ファイル日付 に変換して比較する
    start_file = start + datetime.timedelta(days=1)
    end_file = end + datetime.timedelta(days=1)
    in_range = [i for i, (date, _) in enumerate(dated_files) if start_file <= date <= end_file]
    if not in_range:
        return [], False
    has_predecessor = in_range[0] > 0
    first = max(0, in_range[0] - 1)
    return dated_files[first : in_range[-1] + 1], has_predecessor


def _make_day_saver(
    *,
    output_path: Path,
    start_date: datetime.date,
    existing_dates: list[datetime.date],
    row_keys: list[tuple[str, str]],
) -> tuple[DaySavedCallback, list[tuple[int, int]]]:
    """``compute_counts`` の ``on_day_done`` フックを組み立てる。

    日1日ぶんの突き合わせが終わるたびに呼ばれ、ここまでの累積状態を CSV へ
    書き出す。CSV は少量のデータ（8行 × 数百列で数KB）なので、毎回まるごと
    書き直しても Excel 1ファイルの読み込み（数万行）に比べて無視できる。

    ただし、保存のたびに ``end_date`` まで全部の列を書き出すと、未比較の区間も
    列として出てしまい、``last_date_in_csv`` が本来の位置より先を指してしまう
    （= 再開時に未処理の日を飛ばしてしまう）。そこで、毎回 ``start_date`` から
    「この保存までに比較した最新の日」までの range だけを列にする。

    戻り値は ``(保存コールバック, これまでの保存サイズのリスト)``。``save_sizes``
    には ``write_csv`` が返した ``(行数, 列数)`` が毎回追加される。``run.py``
    の最後で ``save_sizes[-1]`` を使い、「出力しました」を1回だけ出すために
    使う（途中保存のたびに毎回ログを出すと、237ファイルぶんの実行でログが
    埋もれてしまうため）。
    """

    sizes: list[tuple[int, int]] = []

    def _save(
        by_row: ByRow,
        compared_dates: set[datetime.date],
        target_dates_by_month: dict[str, set[datetime.date]] | None = None,
    ) -> None:
        last_compared = max(compared_dates)  # 1日ぶん終わった直後なので必ず空でない
        range_dates = _date_range(start_date, last_compared)
        # この保存時点で既に比較済みの日だけ ``True`` にする。比較済みの日は
        # 「ファイルはあったが件数が0」を ``"0"`` で表す対象になる。
        all_dates = sorted(set(range_dates) | set(existing_dates))
        dates = [(d, d in compared_dates) for d in all_dates]
        sizes.append(write_csv(output_path, by_row, dates, row_keys, target_dates_by_month))

    return _save, sizes
