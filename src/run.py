"""
src/run.py — 処理の本体

フォルダの中で日付が新しい2ファイル（当日・前日）を突き合わせ、
キー列の値の出入りを「延期」「積み上げ」として集計する。

    前日にあって当日にない → 延期（その日付が対象期間の外へ動いた）
    前日になくて当日にある → 積み上げ（対象期間の中へ入ってきた）
"""

import logging
from pathlib import Path

from comken.utils.files import FileFinder, date_in_name

from .exceptions import ComparisonFileNotEnoughError
from .report import write_report
from .settings import Settings
from .source import Record, read_records

logger = logging.getLogger(__name__)

COMPARISON_FILE_COUNT = 2  # 当日と前日
OUTPUT_NAME_FORMAT = "延期積上集計_{before}_{after}.xlsx"


def run(settings: Settings) -> Path:
    """当日と前日を比べて集計ブックを作り、その出力先を返す。

    Args:
        settings: config.ini から読んだ設定一式。

    Raises:
        ComparisonFileNotEnoughError: 比較できるファイルが2つ未満の場合。
    """
    after_path, before_path = _comparison_files(settings)
    logger.info("当日: %s / 前日: %s", after_path.name, before_path.name)

    before = read_records(before_path, settings.layout, settings.criteria)
    after = read_records(after_path, settings.layout, settings.criteria)

    # 差分はキー列の値の有無だけで決まるので、列の値までは比べない
    # （comken の diff_rows は全列を突き合わせるため、何万行では無駄が大きい）
    postponed = [before[key] for key in before.keys() - after.keys()]
    added = [after[key] for key in after.keys() - before.keys()]
    logger.info("積み上げ %d 件 / 延期 %d 件", len(added), len(postponed))
    _log_unchanged(before, after)

    output_path = settings.output_folder / OUTPUT_NAME_FORMAT.format(
        before=_date_text(before_path), after=_date_text(after_path)
    )
    write_report(output_path, added, postponed, settings.criteria, settings.layout)
    return output_path


def _comparison_files(settings: Settings) -> tuple[Path, Path]:
    """フォルダから当日・前日の順で2ファイルを取り出す。"""
    files = FileFinder(settings.input_folder).dated(settings.file_pattern, required=False)
    if len(files) < COMPARISON_FILE_COUNT:
        raise ComparisonFileNotEnoughError(
            str(settings.input_folder), settings.file_pattern, len(files)
        )
    return files[0], files[1]  # dated() は日付の新しい順


def _log_unchanged(before: dict[str, Record], after: dict[str, Record]) -> None:
    """両日にあった件数を記録する（差分が極端に多いときの確認材料にする）。"""
    logger.info("前日・当日の両方にあった件数: %d 件", len(before.keys() & after.keys()))


def _date_text(path: Path) -> str:
    """出力ファイル名に使う、元ファイル名の日付（YYYYMMDD）。"""
    date = date_in_name(path.name)
    return date.strftime("%Y%m%d") if date else path.stem
