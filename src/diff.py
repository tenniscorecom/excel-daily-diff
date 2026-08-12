"""
src/diff.py — 日付順に並んだファイルを、隣り合う日どうしで突き合わせる

「その日に何件延期になり、何件積み上がったか」を1日ぶんずつ出す。
ある日の結果は、その日のファイルと**その直前のファイル**を比べたもの。
土日祝でファイルが飛んでいれば、飛んだ手前のファイルが比較相手になる。
"""

import datetime
import logging
from dataclasses import dataclass
from pathlib import Path

from .settings import Criteria, SourceLayout
from .source import Record, read_records

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class DailyDiff:
    """1日ぶんの差分。"""

    date: datetime.date  # 集計表の横軸になる、後ろ側のファイルの日付
    before_name: str
    after_name: str
    added: list[Record]  # 積み上げ
    postponed: list[Record]  # 延期


def daily_diffs(
    dated_files: list[tuple[datetime.date, Path]],
    layout: SourceLayout,
    criteria: Criteria,
) -> list[DailyDiff]:
    """日付の古い順に並んだファイルを、隣り合う組で突き合わせる。

    Args:
        dated_files: (ファイル名の日付, パス) を日付の古い順に並べたもの。
            先頭は比較相手として読むだけで、結果には出ない。
        layout: シート名と列名。
        criteria: 対象月・種別・状態などの条件。

    Returns:
        2件目以降のファイルぶんの DailyDiff。
    """
    diffs: list[DailyDiff] = []
    previous_records: dict[str, Record] | None = None
    previous_name = ""
    for date, path in dated_files:
        records = read_records(path, layout, criteria)
        if previous_records is not None:
            diffs.append(
                DailyDiff(
                    date=date,
                    before_name=previous_name,
                    after_name=path.name,
                    added=[records[key] for key in records.keys() - previous_records.keys()],
                    postponed=[
                        previous_records[key] for key in previous_records.keys() - records.keys()
                    ],
                )
            )
            logger.debug(
                "%s: 積み上げ %d 件 / 延期 %d 件",
                date,
                len(diffs[-1].added),
                len(diffs[-1].postponed),
            )
        previous_records = records
        previous_name = path.name

    logger.info(
        "%d 日ぶんを比較しました（積み上げ 合計 %d 件 / 延期 合計 %d 件）",
        len(diffs),
        sum(len(d.added) for d in diffs),
        sum(len(d.postponed) for d in diffs),
    )
    return diffs
