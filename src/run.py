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
8 月でも、業務日 1 月の列は 1 月の案件だけを見る（= 1 月時点の積み上げ・
延期を 1 月の案件について見る、という集計の意図に沿う）。23 日以降は翌月も
対象に含める判定は、実行日ではなく業務日で行う。
"""

import datetime
import logging
from pathlib import Path

from comken import config
from comken.core import DateFileFinder, date_in_name, today as _today
from comken.core.files import atomic_write

from src.diff import ByRow, DaySavedCallback, compute_counts
from src.report import COL_TARGET_MONTH, last_date_in_csv, read_existing, write_csv
from src.source import load_rules

logger = logging.getLogger(__name__)

# 集計対象月の境界。この日以降なら翌月も対象に加える
NEXT_MONTH_FROM_DAY = 23

# 出力ファイル名（[REPORT] OUTPUT_FOLDER の下に固定名で置く。日付入りにはしない）
OUTPUT_NAME = "集計.csv"
CONDITIONS_NAME = "集計.conditions.txt"


def run(today: datetime.date | None = None) -> Path:
    """集計を実行し、出力先 CSV のパスを返す。

    Args:
        today: 実行日。省略時は ``comken.core.today()`` を使う。テストで日付を固定したいときに渡す。
    """
    if today is None:
        today = _today()

    plan_prefixes = tuple(config.FILTER.PLAN_PREFIXES)
    kinds = tuple(config.FILTER.KINDS)
    rules = load_rules()
    input_folder = Path(config.FILES.INPUT_FOLDER)
    output_folder = Path(config.REPORT.OUTPUT_FOLDER)
    output_path = output_folder / OUTPUT_NAME
    conditions_path = output_folder / CONDITIONS_NAME

    dated_files = _all_dated_files(input_folder)
    if not dated_files:
        logger.warning("入力フォルダに対象ファイルがありません: %s", input_folder)
        return output_path

    existing = read_existing(output_path)
    current_conditions = _current_conditions_text(plan_prefixes, kinds, rules)
    previous_conditions = _read_text(conditions_path)
    conditions_changed = (
        previous_conditions is not None and previous_conditions != current_conditions
    )
    if conditions_changed:
        logger.info("[条件が変わりました] 過去ぶんは古い条件で数えられているため、全期間を作り直します")
        logger.info("--- 現在の条件 ---\n%s", current_conditions)

    # 下限：実行日の年の1月1日（業務日）。それより古いファイルは対象外
    # （古いファイルはシート構造が違うことがある）
    range_floor = _range_floor(today)
    # 上限：実行日（業務日）。今日より後の日付のファイルは対象外
    range_ceiling = today

    last_date = last_date_in_csv(existing) if existing else None
    if conditions_changed or last_date is None:
        # 全期間を作り直す
        start_date = range_floor
        run_mode = "条件変更による全期間の作り直し" if conditions_changed else "初回"
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

    # 実行開始時の概要ログ（範囲・対象月・読み込み/読み飛ばし件数・比較相手を1か所で示す）
    _log_run_header(
        input_folder=input_folder,
        target_months=target_months_in_range,
        run_mode=run_mode,
        start_date=start_date,
        end_date=end_date,
        dated_files=dated_files,
        targets=targets,
        has_predecessor=has_predecessor,
    )

    # 既存 CSV に残っている日付列（再開時に過去側の列が消えないように）。
    # 毎回保存時に ``start_date`` から「直近で比較が終わった日」までの range と
    # 合併して列見出しを作る（未比較の区間があると再開位置がズレるため）。
    existing_dates = _dates_from_existing(existing)

    # 書き出す行の (対象月, 種別) の組を呼び出し側で組み立てて渡す。
    # 読み込み範囲全体で対象月になる月の行を必ず作り、既存 CSV にある対象月も
    # 残す（古い対象月の行を消さない）。
    existing_months = _months_from_existing(existing)
    months_for_rows = sorted(set(target_months_in_range) | existing_months)
    row_keys = [(f"{year:04d}-{month:02d}", plan)
                for (year, month) in months_for_rows
                for plan in plan_prefixes]

    counts = compute_counts(
        targets,
        _target_months,
        plan_prefixes,
        kinds,
        rules,
        range_start=start_date,
        on_day_done=_make_day_saver(
            output_path=output_path,
            conditions_path=conditions_path,
            current_conditions=current_conditions,
            start_date=start_date,
            existing_dates=existing_dates,
            row_keys=row_keys,
        ),
    )

    logger.info(
        "%s 〜 %s のうち %d 日ぶんを比較しました（対象月: %s）",
        start_date,
        end_date,
        len(counts.compared_dates),
        ", ".join(f"{y:04d}-{m:02d}" for y, m in target_months_in_range),
    )
    return output_path


def _target_months(business_date: datetime.date) -> list[tuple[int, int]]:
    """業務日から対象月を決める。23 日以降なら翌月も加える。

    集計表の横軸である業務日 D に対して、D の月と（23日以降なら）D の翌月を
    対象月として返す。**実行日ではなく業務日**を見る点が重要で、これにより
    業務日 1月の列は 1月の案件だけを対象にし、業務日 1/23 の列は 1月と 2月の
    両方を対象にする（実行日がどちらでも変わらない）。
    """
    months = [(business_date.year, business_date.month)]
    if business_date.day >= NEXT_MONTH_FROM_DAY:
        if business_date.month == 12:
            months.append((business_date.year + 1, 1))
        else:
            months.append((business_date.year, business_date.month + 1))
    return months


def _target_months_in_range(
    start: datetime.date, end: datetime.date
) -> list[tuple[int, int]]:
    """``start`` から ``end`` までの各業務日について対象月を求め、和集合を返す。

    業務日ごとに ``_target_months`` を呼ぶので、月をまたぐ境界や 23 日をまたぐ
    業務日が含まれていれば、自動的に複数の対象月が並ぶ。順序は年月昇順。
    """
    months: set[tuple[int, int]] = set()
    current = start
    while current <= end:
        months.update(_target_months(current))
        current += datetime.timedelta(days=1)
    return sorted(months)


def _date_range(start: datetime.date, end: datetime.date) -> list[datetime.date]:
    """``start`` から ``end`` までの全日付を古い順で返す（1日も飛ばさない）。"""
    return [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]


def _log_run_header(
    *,
    input_folder: Path,
    target_months: list[tuple[int, int]],
    run_mode: str,
    start_date: datetime.date,
    end_date: datetime.date,
    dated_files: list[tuple[datetime.date, Path]],
    targets: list[tuple[datetime.date, Path]],
    has_predecessor: bool,
) -> None:
    """実行開始時の概要ログを出す。

    「入力フォルダ／パターン／対象月／実行モード／読み込み範囲」と、
    「フォルダ全件のうち何件読み、何件読み飛ばしたか」を1セットにして出す。
    範囲外から比較相手として読む1ファイルがあれば、そのファイル名を別行で明示する。

    ``target_months`` は読み込み範囲全体で対象月になりうる月の和集合。
    業務日ごとに対象月が変わるので、開始ログにはその全体を出す
    （業務日ごとの対象月は日次のログに別途出る）。

    ``start_date`` / ``end_date`` は業務日。``targets`` の ``date`` はファイル日付
    （= 業務日 + 1日）なので、件数カウントで ``+1日`` して比較する。
    """
    logger.info("入力フォルダ: %s（パターン: %s）", input_folder, config.FILES.FILE_PATTERN)
    logger.info(
        "対象月: %s",
        ", ".join(f"{y:04d}-{m:02d}" for y, m in target_months),
    )
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


def _range_floor(today: datetime.date) -> datetime.date:
    """読み込み対象の下限 = 実行日の年の1月1日。

    入力フォルダに何年ぶん溜まっていても、この年より古いファイルは開かない
    （古いファイルはシート構造が違うことがあるため）。
    1月に実行したとき、前年12月のファイルは範囲外だが、比較相手として
    直前1ファイルだけ例外的に読まれる。
    """
    return datetime.date(today.year, 1, 1)


def _dates_from_existing(existing: list[dict[str, object]]) -> list[datetime.date]:
    """既存 CSV の1行目から、日付ヘッダとして解釈できたものを抽出する。"""
    if not existing:
        return []
    first_col_index = 3  # 対象月 / 種別 / 判定 の3列のあと
    headers = list(existing[0].keys())[first_col_index:]
    dates: list[datetime.date] = []
    for header in headers:
        try:
            dates.append(datetime.date.fromisoformat(header.strip()))
        except ValueError:
            continue
    return dates


def _months_from_existing(existing: list[dict[str, object]]) -> set[tuple[int, int]]:
    """既存 CSV の「対象月」列を読み取り、``(year, month)`` の集合を返す。

    既存の対象月の行を残すために使う。解釈できない行はスキップする
    （手編集などで壊れた行があっても落とさず、後段の ``write_csv`` が
    ``leftover_rows`` としてそのまま残すので問題ない）。
    """
    months: set[tuple[int, int]] = set()
    for row in existing:
        value = row.get(COL_TARGET_MONTH, "")
        if not value:
            continue
        text = str(value).strip()
        try:
            year_str, month_str = text.split("-")
            months.add((int(year_str), int(month_str)))
        except ValueError:
            continue
    return months


def _all_dated_files(folder: Path) -> list[tuple[datetime.date, Path]]:
    """フォルダ内の対象ファイルを、日付の古い順に並べて返す。

    config.ini の ``FILE_PATTERN`` は ``一覧_*.xlsx`` のように ``*`` を含めて
    書く。``DateFileFinder.dated`` は ``*`` をワイルドカード扱いせず文字どおりの
    前方一致で探すので、ここで ``*`` を落として渡す。
    """
    prefix = str(config.FILES.FILE_PATTERN).replace("*", "")
    found = DateFileFinder(folder).dated(prefix)
    dated = sorted(
        (date, path) for path in found if (date := date_in_name(path.name)) is not None
    )
    return dated


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


def _current_conditions_text(
    plan_prefixes: tuple[str, ...],
    kinds: tuple[str, ...],
    rules: tuple,
) -> str:
    """いま使っている条件を人が読めるテキストにする。"""
    lines = [
        f"SHEET_NAME = {config.SOURCE.SHEET_NAME}",
        f"HEADER_ROW = {config.SOURCE.HEADER_ROW}",
        f"KEY_COLUMN = {config.SOURCE.KEY_COLUMN}",
        f"DATE_COLUMN = {config.SOURCE.DATE_COLUMN}",
        f"PLAN_COLUMN = {config.SOURCE.PLAN_COLUMN}",
        f"KIND_COLUMN = {config.SOURCE.KIND_COLUMN}",
        f"PLAN_PREFIXES = {list(plan_prefixes)}",
        f"KINDS = {list(kinds)}",
    ]
    for rule in rules:
        kind = "含む" if rule.is_contains else "完全一致"
        action = "外す" if rule.is_exclude else "残す"
        lines.append(
            f"RULE: [{rule.column}] {','.join(rule.words)} ({kind}/{action})"
        )
    return "\n".join(lines)


def _read_text(path: Path) -> str | None:
    if not path.exists():
        return None
    return path.read_text(encoding="utf-8")


def _write_text(path: Path, content: str) -> None:
    """テキストファイルを書き込む。途中保存と組み合わせるため、原子的書き出しにする。

    ``os.replace`` で一括置換するため、書き出し中に落ちても書きかけの状態で
    残らない（部分的に書き込まれた状態だと、次回起動時に内容が壊れたように見える）。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with atomic_write(path) as tmp_path:
        tmp_path.write_text(content, encoding="utf-8")


def _make_day_saver(
    *,
    output_path: Path,
    conditions_path: Path,
    current_conditions: str,
    start_date: datetime.date,
    existing_dates: list[datetime.date],
    row_keys: list[tuple[str, str]],
) -> DaySavedCallback:
    """``compute_counts`` の ``on_day_done`` フックを組み立てる。

    日1日ぶんの突き合わせが終わるたびに呼ばれ、ここまでの累積状態を CSV へ
    書き出す。条件ファイルは**最初の保存と同時に**1回だけ書く（途中で落ちても、
    次回が「条件は変わっていない」と判定できるように）。

    CSV は少量のデータ（8行 × 数百列で数KB）なので、毎回まるごと書き直しても
    Excel 1ファイルの読み込み（数万行）に比べて無視できる。N 件ごとに保存する
    ような間隔の定数は、ここでは意図的に持たない。

    ただし、保存のたびに ``end_date`` まで全部の列を書き出すと、未比較の区間も
    列として出てしまい、``last_date_in_csv`` が本来の位置より先を指してしまう
    （= 再開時に未処理の日を飛ばしてしまう）。そこで、毎回 ``start_date`` から
    「この保存までに比較した最新の日」までの range だけを列にする。
    """
    conditions_written = False

    def _save(
        by_row: ByRow,
        compared_dates: set[datetime.date],
        _date: datetime.date,
        target_dates_by_month: dict[tuple[int, int], set[datetime.date]] | None = None,
    ) -> None:
        nonlocal conditions_written
        last_compared = max(compared_dates)  # 1日ぶん終わった直後なので必ず空でない
        range_dates = _date_range(start_date, last_compared)
        # この保存時点で既に比較済みの日だけ ``True`` にする。比較済みの日は
        # 「ファイルはあったが件数が0」を ``"0"`` で表す対象になる。
        all_dates = sorted(set(range_dates) | set(existing_dates))
        dates = [(d, d in compared_dates) for d in all_dates]
        write_csv(output_path, by_row, dates, row_keys, target_dates_by_month)
        if not conditions_written:
            _write_text(conditions_path, current_conditions)
            conditions_written = True

    return _save
