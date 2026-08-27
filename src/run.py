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

出力CSV の行は「当月/来月ラベル × 種別（``[FILTER] PLAN_PREFIXES`` の設定順）
× 判定（積み上げ/延期）」の固定構成。``PLAN_PREFIXES`` が2件のとき 8 行で、
件数が変われば行数も変わる（= 「常に 8 行」ではない）。
"""

import datetime
import functools
import logging
from pathlib import Path

from comken import config
from comken.core import DateFileFinder, date_in_name, today

from src.diff import ByRow, DaySavedCallback, compute_counts
from src.report import ROW_LABELS, last_date_in_csv, read_existing, write_csv
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
    # comken の ``Config`` は絶対パスしか自動で ``Path`` に変換しないため、
    # 相対パスが書かれていると ``/`` 演算子で ``TypeError`` になる。``Path()``
    # で明示的に包むことで絶対・相対どちらの設定でも動く。
    input_folder = Path(config.FILES.INPUT_FOLDER)
    output_folder = Path(config.REPORT.OUTPUT_FOLDER)
    output_path = output_folder / OUTPUT_NAME

    # ローリングモード判定（[FILES] ROLLING_WINDOW_DAYS）。キーが無い／読み取れない
    # 場合は None（=年次累積モード。既存挙動と完全同一）。
    rolling_window_days = _resolve_rolling_window_days()

    dated_files = _all_dated_files(input_folder)
    if not dated_files:
        logger.warning("入力フォルダに対象ファイルがありません: %s", input_folder)
        return output_path

    existing = read_existing(output_path)

    # 下限：年次累積モードは実行日の年の1月1日（業務日）、ローリングモードは
    # 「今日 - (N-1)日」。ローリングモードでは古いファイルはそもそも対象外。
    range_floor = _range_floor(today_date, rolling_window_days)
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
    # 「次に計算すべき業務日」が「計算可能な上限」を超えているとき
    # （= ``_files_in_range`` が必ず空リストを返す状態）は、先に専用ログを出して
    # 早期リターンする。``start_date > end_date`` を満たすとき ``start_file > end_file``
    # となるので ``_files_in_range`` は空確定。ここで先に拾うことで、
    # 後続の「範囲内ファイル無し」分岐が誤って警告を出さずに済む。
    if start_date > end_date:
        if last_date is not None:
            # 日次実行で日常的に起きる「追いついた」状態。warning ではなく info。
            # 次の業務日を計算するには start_date + 1日 付けのファイル（= 業務日
            # start_date 終了時点のデータ）が要るが、まだ届いていない。
            logger.info(
                "既に %s 終了時点まで計算済みです。次に %s 終了時点を計算するには "
                "%s 付けのファイルが必要です（まだ入力フォルダにありません）",
                last_date,
                start_date,
                (start_date + datetime.timedelta(days=1)).isoformat(),
            )
        else:
            # 初回モード（last_date が None）で range_floor 以降のファイルが
            # 1 つも無いという稀なケース。元の意味に近い警告のまま出す。
            logger.warning(
                "入力フォルダに %s 以降のファイルが見当たりません: %s",
                range_floor,
                input_folder,
            )
        return output_path
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

    # 業務日基準の対象月（実際の暦月文字列）。読み込み範囲全体で登場する
    # 対象月の和集合を取る。実行日基準の単一グループではないので、開始ログは
    # 「範囲内で出現しうる対象月」の全体を出す（業務日ごとのログにはその日に
    # 使った対象月が出る）。
    target_months_in_range = _target_months_in_range(
        start_date, end_date, rolling_window_days=rolling_window_days
    )

    # 実行開始時の概要ログ。処理の進行に合わせて直書きする（途中でどこまで
    # 進んでいたかがログから追えるように）。
    logger.info("入力フォルダ: %s（パターン: %s）", input_folder, config.FILES.FILE_PATTERN)
    logger.info("対象月: %s", ", ".join(target_months_in_range))
    logger.info("実行モード: %s", run_mode)
    logger.info("読み込み範囲（終了日）: %s 〜 %s", start_date, end_date)
    logger.info(
        "読み込んだファイルの最新日付: %s（このファイルの中身は %s 終了時点のデータです）",
        targets[-1][0],
        end_date,
    )
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
            f"。{start_date}（終了日）の比較相手として、"
            f"範囲外から 1 ファイル追加で読みます（{predecessor_name}）"
        )
    else:
        summary += "（比較相手はありません）"
    logger.info("%s", summary)

    # 既存 CSV に残っている日付列（再開時に過去側の列が消えないように）。
    # 毎回保存時に ``start_date`` から「直近で比較が終わった日」までの range と
    # 合併して列見出しを作る（未比較の区間があると再開位置がズレるため）。
    existing_dates = _dates_from_existing(existing)

    # 書き出す行は「当月/来月のラベル × ``PLAN_PREFIXES`` の種別順」の
    # 直積（× {積み上げ, 延期} で 2 倍）。``PLAN_PREFIXES`` が既定の
    # ``[標準, 上位]`` のときは 2 × 2 × 2 = 8 行になる。
    # 行の対象月は業務日基準の固定ラベル（年別累積/ローリングで対象月が変わる
    # ことは無い）なので、対象月ごとに増減する複雑さは不要のまま。
    row_keys = [(label, plan) for label in ROW_LABELS for plan in plan_prefixes]

    save_callback, save_sizes = _make_day_saver(
        output_path=output_path,
        start_date=start_date,
        existing_dates=existing_dates,
        row_keys=row_keys,
        window_floor=range_floor if rolling_window_days is not None else None,
    )
    # 1日ぶん終わるたびに CSV を途中保存することを最初に1回だけ伝える。
    # 長い処理で「今どこまで保存されているのか」がログから追えるようにするための
    # 既存の方針と整合させる。途中で落ちてもここまでがファイルに残っている。
    logger.info("1日ぶんごとに保存しながら進めます")

    counts = compute_counts(
        targets,
        functools.partial(_target_months, rolling_window_days=rolling_window_days)
        if rolling_window_days is not None
        else _target_months,
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


def _range_floor(
    today_date: datetime.date,
    rolling_window_days: int | None = None,
) -> datetime.date:
    """読み込み範囲の下限（業務日）。

    引数:
        ``rolling_window_days``: ``None``（既定）のときは実行日の年の1月1日を返す
        （年次累積モード。1月実行時でも同年の1月1日。前年の12月1日にはしない）。
        整数を渡したとき（ローリングモード）は **今日を含めて N 日ぶん**の
        暦日になるよう、今日 - (N - 1) 日 を返す。N=1 なら今日だけ、N=7 なら
        今日から6日前まで。

    年次累積モードでは古いファイルはシート構造が違うことが多いため、範囲外との
    比較相手に使う1ファイルを除いて、今年ぶんだけを読む運用が前提。
    ローリングモードでは日々サッと確認する用途のため、窓の幅を「直近 N 日」に
    限定する（1年ぶん全部を読み直す必要が無い）。
    """
    if rolling_window_days is None:
        return datetime.date(today_date.year, 1, 1)
    return today_date - datetime.timedelta(days=rolling_window_days - 1)


def _resolve_rolling_window_days() -> int | None:
    """``config.FILES.ROLLING_WINDOW_DAYS`` を安全に読む。キーが無ければ ``None``。

    comken の ``_SectionNamespace.__getattr__`` は未定義キーで
    ``ConfigKeyNotFoundError``（``AttributeError`` のサブクラス）を投げるため、
    ``hasattr`` で存在判定する。``_parse_value`` が整数を ``int`` に変換するので、
    設定されていれば ``int`` が返る（``True`` / ``False``/リスト/Path 等ではない）。
    """
    if not hasattr(config.FILES, "ROLLING_WINDOW_DAYS"):
        return None
    value = config.FILES.ROLLING_WINDOW_DAYS
    if isinstance(value, bool):
        # ``True`` / ``False`` も ``int`` のサブクラスで ``isinstance(x, int)`` が
        # True になる。``ROLLING_WINDOW_DAYS`` を真偽値として書かれたケースは
        # 設定ミスなので、誤って巨大な数値として扱わないよう明示的に拒否する
        raise TypeError(
            "ROLLING_WINDOW_DAYS には整数を指定してください"
            f"（bool は不可）: {value!r}"
        )
    if not isinstance(value, int):
        raise TypeError(
            "ROLLING_WINDOW_DAYS には整数を指定してください"
            f"（現在: {type(value).__name__}）"
        )
    if value < 1:
        raise ValueError(
            f"ROLLING_WINDOW_DAYS には 1 以上の整数を指定してください（現在: {value}）"
        )
    return value


def _target_months(
    business_date: datetime.date,
    *,
    rolling_window_days: int | None = None,
) -> list[str]:
    """業務日から対象月を決める。

    既定（``rolling_window_days=None``）は業務日の属する月だけを返す（既存の
    年次累積モードの挙動、完全不変）。

    ``rolling_window_days`` を渡したとき（ローリングモード）は **無条件で**
    当月と翌月の両方を対象月に加える。業務日が月初でも月末でも、月内のどの
    日であっても、戻り値は常に ``[当月, 翌月]`` の2要素になる（= 旧来の
    「月末 N 日前からの lookahead」条件は廃止）。月単位のラベルに先取りで
    件数を入れたいのはローリングモード固有の要件で、年次累積モードでは
    翌月を先取りしない（「来月」行は構造上存在するが空のまま）。

    月末が 12 月のときは翌月が翌年 1 月になるため、ラベル計算は年跨ぎも
    正しく扱う。

    戻り値は **必ず業務日自身の月が 0 番目、翌月が 1 番目** という順序を保つ
    （``src/diff.py`` の ``_label_for_month_index`` がこの順序に依存して
    「当月」「来月」のラベルへ変換する）。
    """
    months = [f"{business_date.year:04d}-{business_date.month:02d}"]
    if rolling_window_days is not None:
        if business_date.month == 12:
            next_year = business_date.year + 1
            next_month = 1
        else:
            next_year = business_date.year
            next_month = business_date.month + 1
        months.append(f"{next_year:04d}-{next_month:02d}")
    return months


def _target_months_in_range(
    start: datetime.date,
    end: datetime.date,
    *,
    rolling_window_days: int | None = None,
) -> list[str]:
    """``start`` から ``end`` までの各業務日について対象月を求め、和集合を返す。

    業務日ごとに ``_target_months`` を呼ぶので、月をまたぐ境界があれば
    自動的に複数の対象月が並ぶ。順序は年月昇順。
    ローリングモードでは ``_target_months`` 側にキーワード引数を渡して、
    月末近くの業務日では翌月も拾う（``_target_months`` の docstring 参照）。
    """
    months: set[str] = set()
    current = start
    while current <= end:
        months.update(
            _target_months(current, rolling_window_days=rolling_window_days)
        )
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


def _all_dated_files(folder: Path) -> list[tuple[datetime.date, Path]]:
    """フォルダ内の対象ファイルを、日付の古い順に並べて返す。

    config.ini の ``FILE_PATTERN`` は ``一覧_*.xlsx`` のように ``*`` を含めて
    書く。``DateFileFinder.dated`` は ``*`` をワイルドカード扱いせず文字どおりの
    前方一致で探すので、ここで ``*`` を落として渡す。
    """
    prefix = config.FILES.FILE_PATTERN.replace("*", "")
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
    window_floor: datetime.date | None = None,
) -> tuple[DaySavedCallback, list[tuple[int, int]]]:
    """``compute_counts`` の ``on_day_done`` フックを組み立てる。

    日1日ぶんの突き合わせが終わるたびに呼ばれ、ここまでの累積状態を CSV へ
    書き出す。CSV は少量のデータ（行数 ＝ ``len(row_keys) * 2`` × 数百列で数KB）
    なので、毎回まるごと書き直しても Excel 1ファイルの読み込み（数万行）に
    比べて無視できる。

    ただし、保存のたびに ``end_date`` まで全部の列を書き出すと、未比較の区間も
    列として出てしまい、``last_date_in_csv`` が本来の位置より先を指してしまう
    （= 再開時に未処理の日を飛ばしてしまう）。そこで、毎回 ``start_date`` から
    「この保存までに比較した最新の日」までの range だけを列にする。

    ローリングモードでは「直近 N 日の窓」があり、``window_floor`` が指定された
    ときは ``existing_dates`` のうち窓の外にある日付を捨ててから range と合併する
    （= 窓の外に出た古い列を毎回保存で消す）。行は対象月が業務日基準の固定
    ラベルなので、窓の外に出た行を捨てるような仕組みは要らない。
    年次累積モード（既定）は ``window_floor=None`` で、既存挙動と完全同一。

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
        # ローリングモードでは窓の外の日付は捨てる
        base_existing = (
            [d for d in existing_dates if window_floor is None or d >= window_floor]
        )
        # この保存時点で既に比較済みの日だけ ``True`` にする。比較済みの日は
        # 「ファイルはあったが件数が0」を ``"0"`` で表す対象になる。
        all_dates = sorted(set(range_dates) | set(base_existing))
        dates = [(d, d in compared_dates) for d in all_dates]
        sizes.append(write_csv(
            output_path, by_row, dates, row_keys, target_dates_by_month,
        ))

    return _save, sizes
