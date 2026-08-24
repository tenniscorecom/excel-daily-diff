"""
main.py — エントリポイント

このプロジェクトの入口。`python main.py` で実行できる（非エンジニアは 実行.bat をダブルクリック）。

処理の本体は src/ 以下に書き、ここでは「実行 → エラーの受け止め」だけを行う。
実行は社内 RPA 基盤を通す（設定の初期化と時間計測は基盤側が行う）。
"""

import logging

from comken.exceptions import ComkenError
from comken.internal.rpa import backoffice  # イントラネットのツールなら intranet に変える

from src.run import run
from src.settings import load_settings

logger = logging.getLogger(__name__)

# 社内 RPA 基盤へ渡すプロジェクト名。ログの識別に使われる
PROJECT_NAME = "延期積上集計"


def main() -> None:
    run(load_settings())


if __name__ == "__main__":
    # ログの設定は社内の共通ライブラリ側で行う。ここでは logging をそのまま使う
    # 動作確認だけしたいときは保存・送信をスキップできる:
    #   from comken import dry_run
    #   with dry_run():
    #       backoffice(main, PROJECT_NAME)
    try:
        # main を直接呼ばず基盤に渡す。基盤が設定の初期化と時間計測をしてから main を呼ぶ
        backoffice(main, PROJECT_NAME)
    except ComkenError as e:
        # comken のエラーはメッセージに対処法が入っている（docs/ERRORS.md も参照）
        logger.error("処理を中断しました: %s", e)
        raise
    except Exception:
        logger.error("予期しないエラーが発生しました", exc_info=True)
        raise
