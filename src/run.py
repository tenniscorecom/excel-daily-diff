"""
src/run.py — 処理の本体

実行日→対象月、既存 CSV→期間の始まり、入力フォルダ→期間の終わり を組み合わせて
集計範囲を決め、日ごとの延期・積み上げを集計する。

日付が2種類出てくるので混同しないこと。
    ファイル名の日付 … 集計表の横軸。いつ時点の一覧かを表す
    案件の日付      … 絞り込みにだけ使う。対象月に入っているかを見る
"""

import datetime
import logging
from pathlib import Path

from comken import config
from comken.core import DateFileFinder, date_in_name, today as _today

from src.diff import compute_counts
from src.report import last_date_in_csv, read_existing, write_csv
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
    target_months = _target_months(today)
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

    last_date = last_date_in_csv(existing) if existing else None
    if conditions_changed or last_date is None:
        # 全期間を作り直す
        start_date = dated_files[0][0]
    else:
        # 増分計算（既存 CSV の最後の日付の次の日から）
        start_date = last_date + datetime.timedelta(days=1)

    end_date = dated_files[-1][0]
    targets = _files_in_range(dated_files, start_date, end_date)

    counts = compute_counts(
        targets,
        target_months,
        plan_prefixes,
        kinds,
        rules,
    )

    # 列として並べる全日付。開始日から終了日まで1日も飛ばさず、
    # 既存 CSV の日付列も残す（増分計算で以前の列が消えないように）。
    dates = _build_dates(start_date, end_date, existing, counts.compared_dates)

    # 書き出す行の (対象月, 種別) の組を呼び出し側で組み立てて渡す。
    row_keys = [(f"{month[0]:04d}-{month[1]:02d}", plan)
                for month in target_months
                for plan in plan_prefixes]

    write_csv(
        output_path,
        counts.by_row,
        dates,
        row_keys,
    )
    _write_text(conditions_path, current_conditions)
    logger.info(
        "%s 〜 %s のうち %d 日ぶんを比較しました（対象月: %s）",
        start_date,
        end_date,
        len(counts.compared_dates),
        ", ".join(f"{y:04d}-{m:02d}" for y, m in target_months),
    )
    return output_path


def _target_months(today: datetime.date) -> list[tuple[int, int]]:
    """実行日から対象月を決める。23 日以降なら翌月も加える。"""
    months = [(today.year, today.month)]
    if today.day >= NEXT_MONTH_FROM_DAY:
        if today.month == 12:
            months.append((today.year + 1, 1))
        else:
            months.append((today.year, today.month + 1))
    return months


def _date_range(start: datetime.date, end: datetime.date) -> list[datetime.date]:
    """``start`` から ``end`` までの全日付を古い順で返す（1日も飛ばさない）。"""
    return [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]


def _build_dates(
    start: datetime.date,
    end: datetime.date,
    existing: list[dict[str, object]],
    compared_dates: set[datetime.date],
) -> list[tuple[datetime.date, bool]]:
    """列として並べる日付を作る。

    開始日から終了日まで1日も飛ばさず、既存 CSV の日付列も追加する。
    今日比較した日は ``True``（``"0"`` を入れる対象）、履歴は ``False``
    （既存値があればそのまま、空なら空セル）。
    """
    range_dates = _date_range(start, end)
    existing_dates = _dates_from_existing(existing)
    all_dates = sorted(set(range_dates) | set(existing_dates))
    return [(d, d in compared_dates) for d in all_dates]


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
) -> list[tuple[datetime.date, Path]]:
    """期間内 + その1つ前のファイルを取り出して返す（最初の日の比較相手）。"""
    in_range = [i for i, (date, _) in enumerate(dated_files) if start <= date <= end]
    if not in_range:
        return []
    first = max(0, in_range[0] - 1)
    return dated_files[first : in_range[-1] + 1]


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
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")