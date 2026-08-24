"""
src/settings.py — config.ini の値を読み、処理で使う形に整える

config.ini を読むのはこのファイルだけにする。列名や日付の書き方を変えたくなったとき、
直す場所がここに集まる。
"""

import datetime
import logging
from dataclasses import dataclass
from pathlib import Path

from comken import config
from comken.exceptions import ConfigSectionNotFoundError

from src.exceptions import (
    InvalidDateSettingError,
    InvalidFilePatternError,
    InvalidMonthSettingError,
)

logger = logging.getLogger(__name__)

CONFIG_DATE_FORMAT = "%Y-%m-%d"
CONFIG_MONTH_FORMAT = "%Y-%m"

# あとから絞り込みを足すためのセクション。書かれていなければ条件なしとして扱う
SECTION_INCLUDE_CONTAINS = "INCLUDE_CONTAINS"
SECTION_EXCLUDE_CONTAINS = "EXCLUDE_CONTAINS"
SECTION_INCLUDE_EXACT = "INCLUDE_EXACT"
SECTION_EXCLUDE_EXACT = "EXCLUDE_EXACT"


@dataclass(frozen=True)
class FilePattern:
    """config.ini の FILE_PATTERN（``一覧_*.xlsx`` のようなワイルドカード表記）を
    ``DateFileFinder.dated()`` が受け取れる ``prefix`` と ``extension`` に分けたもの。"""

    prefix: str
    extension: str


def split_file_pattern(pattern: str) -> FilePattern:
    """``FILE_PATTERN`` を ``DateFileFinder.dated()`` 用に分割する。

    ``*`` の手前を ``prefix``、最後の ``.`` 以降を ``extension`` として取り出す。
    ``*`` が無いパターンや ``*`` が複数あるパターンは対応外（docstring に明記する）。
    """
    star_index = pattern.find("*")
    if star_index < 0:
        raise InvalidFilePatternError(
            pattern,
            "ワイルドカード '*' が含まれていません。",
        )
    if pattern.count("*") > 1:
        raise InvalidFilePatternError(
            pattern,
            "ワイルドカード '*' が複数含まれています。'*' は1つだけ使えます。",
        )
    head = pattern[:star_index]
    dot_index = pattern.rfind(".")
    if dot_index < star_index:
        # '*' の後に '.' が無いケースは拡張子の指定として読めない
        raise InvalidFilePatternError(
            pattern,
            "'*' の後に拡張子（.xlsx など）がありません。",
        )
    return FilePattern(prefix=head, extension=pattern[dot_index + 1 :])


@dataclass(frozen=True)
class SourceLayout:
    """読み取り元 Excel のシートと列名（config.ini の [SOURCE]）。"""

    sheet_name: str
    header_row: int
    key_column: str
    date_column: str
    plan_column: str
    kind_column: str


@dataclass(frozen=True)
class ColumnRule:
    """1つの列に対する、あとから足した絞り込み条件。

    例:「地域」に「離島」が含まれる行を外す → column="地域", words=("離島",),
    is_contains=True, is_exclude=True
    """

    column: str
    words: tuple[str, ...]
    is_contains: bool  # True=語を含むか / False=語と完全に一致するか
    is_exclude: bool  # True=当てはまる行を外す / False=当てはまる行だけ残す


@dataclass(frozen=True)
class Criteria:
    """1行を集計対象にするかの条件（config.ini の [FILTER] と、あとから足したセクション）。

    日付列は「対象月に入っているか」だけを見る。集計表の横軸になる日付とは別物で、
    こちらは案件そのものの日付（工事の日など）を指す。
    """

    target_year: int
    target_month: int
    plan_prefixes: tuple[str, ...]
    kinds: tuple[str, ...]
    rules: tuple[ColumnRule, ...] = ()


@dataclass(frozen=True)
class Settings:
    """このツールが使う設定一式。"""

    input_folder: Path
    file_pattern: FilePattern  # prefix / extension に分割済み
    start_date: datetime.date  # 集計表の横軸（ファイル名の日付）の始まり
    end_date: datetime.date  # 同じく終わり
    output_folder: Path
    layout: SourceLayout
    criteria: Criteria


def load_settings() -> Settings:
    """config.ini を読んで Settings に詰める。

    Raises:
        InvalidDateSettingError: [FILES] の日付が YYYY-MM-DD で書かれていない場合。
        InvalidMonthSettingError: [FILTER] TARGET_MONTH が YYYY-MM で書かれていない場合。
        InvalidFilePatternError: [FILES] FILE_PATTERN の書き方が split_file_pattern() で扱えない場合。
    """
    layout = SourceLayout(
        sheet_name=str(config.SOURCE.SHEET_NAME),
        header_row=int(config.SOURCE.HEADER_ROW),
        key_column=str(config.SOURCE.KEY_COLUMN),
        date_column=str(config.SOURCE.DATE_COLUMN),
        plan_column=str(config.SOURCE.PLAN_COLUMN),
        kind_column=str(config.SOURCE.KIND_COLUMN),
    )
    year, month = _to_year_month(config.FILTER.TARGET_MONTH)
    criteria = Criteria(
        target_year=year,
        target_month=month,
        plan_prefixes=tuple(config.FILTER.PLAN_PREFIXES),
        kinds=tuple(config.FILTER.KINDS),
        rules=_load_rules(),
    )
    return Settings(
        input_folder=Path(config.FILES.INPUT_FOLDER),
        file_pattern=split_file_pattern(str(config.FILES.FILE_PATTERN)),
        start_date=_to_date("START_DATE", config.FILES.START_DATE),
        end_date=_to_date("END_DATE", config.FILES.END_DATE),
        output_folder=Path(config.REPORT.OUTPUT_FOLDER),
        layout=layout,
        criteria=criteria,
    )


def _load_rules() -> tuple[ColumnRule, ...]:
    """あとから足した絞り込みを、config.ini の4セクションから集める。

    書かれていないセクションは条件なしとして飛ばすので、使わないセクションは
    config.ini から消してよい。
    """
    rules: list[ColumnRule] = []
    rules += _rules_in(SECTION_INCLUDE_CONTAINS, is_contains=True, is_exclude=False)
    rules += _rules_in(SECTION_EXCLUDE_CONTAINS, is_contains=True, is_exclude=True)
    rules += _rules_in(SECTION_INCLUDE_EXACT, is_contains=False, is_exclude=False)
    rules += _rules_in(SECTION_EXCLUDE_EXACT, is_contains=False, is_exclude=True)
    for rule in rules:
        logger.info(
            "追加の絞り込み: %s が %s %s",
            rule.column,
            "／".join(rule.words),
            _rule_label(rule),
        )
    return tuple(rules)


def _rules_in(section: str, is_contains: bool, is_exclude: bool) -> list[ColumnRule]:
    """1セクション分の「列名 = 語, 語」を ColumnRule に変える。"""
    try:
        values = vars(getattr(config, section))
    except ConfigSectionNotFoundError:
        return []  # 任意のセクションなので、無ければ条件なし
    rules = []
    for column, value in values.items():
        # 新 comken のセクション名前空間は ``_section`` / ``_keys`` / ``_path`` を
        # 内部用に持つので、絞り込みの列名候補からは外す
        if column.startswith("_"):
            continue
        words = _to_words(value)
        if words:  # 値が空の行は「まだ書いていない」とみなす
            rules.append(ColumnRule(column, words, is_contains, is_exclude))
    return rules


def _to_words(value: object) -> tuple[str, ...]:
    """config.ini の値を語のタプルにする（1語でも [a, b] 形式でも受ける）。"""
    items = value if isinstance(value, list) else str(value).split(",")
    return tuple(str(item).strip() for item in items if str(item).strip())


def _rule_label(rule: ColumnRule) -> str:
    """ログに出す条件の説明。"""
    match = "を含む" if rule.is_contains else "と一致する"
    return f"{match}行だけ残す" if not rule.is_exclude else f"{match}行を外す"


def _to_date(key: str, value: object) -> datetime.date:
    """config.ini に書かれた日付を date にする。"""
    try:
        return datetime.datetime.strptime(str(value), CONFIG_DATE_FORMAT).date()  # noqa: DTZ007
    except ValueError as e:
        raise InvalidDateSettingError(key, str(value)) from e


def _to_year_month(value: object) -> tuple[int, int]:
    """config.ini に書かれた対象月（2026-08）を年と月にする。"""
    text = str(value)
    try:
        parsed = datetime.datetime.strptime(text, CONFIG_MONTH_FORMAT)  # noqa: DTZ007
    except ValueError as e:
        raise InvalidMonthSettingError(text) from e
    if parsed.strftime(CONFIG_MONTH_FORMAT) != text:
        raise InvalidMonthSettingError(text)
    return parsed.year, parsed.month
