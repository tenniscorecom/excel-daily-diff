"""
src/exceptions.py — このツール固有のエラー

main.py が ComkenError でまとめて受け取れるよう、comken の例外体系にぶら下げる。
メッセージには「何が・どこで・どうすればよいか」を書く（docs/ERRORS.md も更新すること）。
"""

from comken.exceptions import ComkenError


class InvalidDateSettingError(ComkenError):
    """config.ini の日付設定が YYYY-MM-DD で書かれていない場合。"""

    def __init__(self, key: str, value: str) -> None:
        super().__init__(
            f"config.ini の [FILES] {key} が日付として読めません: {value}\n"
            "2026-04-21 のように「年-月-日」の形で書いてください。"
        )


class InvalidMonthSettingError(ComkenError):
    """config.ini の対象月が YYYY-MM で書かれていない場合。"""

    def __init__(self, value: str) -> None:
        super().__init__(
            f"config.ini の [FILTER] TARGET_MONTH が月として読めません: {value}\n"
            "2026-08 のように「年-月」の形で書いてください。"
        )


class ComparisonFileNotEnoughError(ComkenError):
    """比べる相手の前日ファイルが見つからない場合。"""

    def __init__(self, folder: str, pattern: str, found: int) -> None:
        super().__init__(
            f"比べるファイルが足りません（{found}件）: {folder}\\{pattern}\n"
            "前日と当日の2ファイルが必要です。フォルダと、ファイル名に日付が"
            "入っているか（一覧_20260812.xlsx のような形か）を確認してください。"
        )


class InvalidFilePatternError(ComkenError):
    """config.ini の FILE_PATTERN の書き方が ``split_file_pattern()`` で扱えない場合。"""

    def __init__(self, pattern: str, reason: str) -> None:
        super().__init__(
            f"config.ini の [FILES] FILE_PATTERN が読み取れません: {pattern}\n"
            f"{reason}\n"
            "例: 一覧_*.xlsx のように、日付の前後に固定部分が残る形（'*' は1つ、"
            "'*' の後に拡張子）で書いてください。"
        )
