"""
src/run.py — 処理の本体

フォルダの中のファイルを日付順に並べ、隣り合う日どうしを順に突き合わせて、
日ごとの延期・積み上げを集計表にする。

    前日にあって当日にない → 延期（その案件が対象月から外れた）
    前日になくて当日にある → 積み上げ（対象月の中へ入ってきた）

日付が2種類出てくるので混同しないこと。
    ファイル名の日付 … 集計表の横軸。いつ時点の一覧かを表す（[FILES] START_DATE〜END_DATE）
    案件の日付      … 絞り込みにだけ使う。対象月に入っているかを見る（[FILTER] TARGET_MONTH）
"""

import datetime
import logging
from pathlib import Path

from comken.core import FileFinder, date_in_name

from src.diff import daily_diffs
from src.exceptions import ComparisonFileNotEnoughError
from src.report import write_report
from src.settings import Settings

logger = logging.getLogger(__name__)

OUTPUT_NAME_FORMAT = "延期積上集計_{start}_{end}.xlsx"
OUTPUT_DATE_FORMAT = "%Y%m%d"


def run(settings: Settings) -> Path:
    """横軸の期間ぶんを日ごとに比較して集計ブックを作り、その出力先を返す。

    Args:
        settings: config.ini から読んだ設定一式。

    Raises:
        ComparisonFileNotEnoughError: 比較できるファイルが足りない場合。
    """
    targets = _target_files(settings)
    logger.info(
        "%s 〜 %s のファイル %d 件を読みます（先頭 %s は比較相手）",
        settings.start_date,
        settings.end_date,
        len(targets),
        targets[0][1].name,
    )

    diffs = daily_diffs(targets, settings.layout, settings.criteria)

    output_path = settings.output_folder / OUTPUT_NAME_FORMAT.format(
        start=settings.start_date.strftime(OUTPUT_DATE_FORMAT),
        end=settings.end_date.strftime(OUTPUT_DATE_FORMAT),
    )
    write_report(output_path, diffs, settings)
    return output_path


def _target_files(settings: Settings) -> list[tuple[datetime.date, Path]]:
    """横軸の期間に入るファイルと、その1つ前（期間の初日の比較相手）を古い順で返す。

    期間の外のファイルまで読むと数万行 × 日数ぶん無駄になるので、必要な範囲だけに絞る。
    """
    found = FileFinder(settings.input_folder).dated(settings.file_pattern, required=False)
    dated_files = sorted(
        (date, path) for path in found if (date := date_in_name(path.name)) is not None
    )
    in_range = [
        i
        for i, (date, _) in enumerate(dated_files)
        if settings.start_date <= date <= settings.end_date
    ]
    if not in_range:
        raise ComparisonFileNotEnoughError(
            str(settings.input_folder), settings.file_pattern, len(dated_files)
        )
    # 期間の初日ぶんを出すには、その手前のファイルが要る
    first = max(0, in_range[0] - 1)
    targets = dated_files[first : in_range[-1] + 1]
    if len(targets) < 2:
        raise ComparisonFileNotEnoughError(
            str(settings.input_folder), settings.file_pattern, len(targets)
        )
    return targets
