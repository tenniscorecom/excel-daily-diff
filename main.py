"""
main.py — エントリポイント

このプロジェクトの入口。`python main.py` で実行できる（非エンジニアは 実行.bat をダブルクリック）。

処理の本体は src/ 以下に書き、ここでは「ログの設定 → run → エラーの受け止め」だけを行う。
RPA 基盤には依存せず、ローカル実行用のロガーを使う。
"""

import logging

from comken.core.logger import setup_local_logging
from comken.exceptions import ComkenError

from src.run import run

logger = logging.getLogger(__name__)

# ログ設定（ローカル実行用）。コンソールと logs/ フォルダへのファイルに出力する
setup_local_logging()


def main() -> None:
    run()


if __name__ == "__main__":
    try:
        main()
    except ComkenError as e:
        # comken のエラーはメッセージに対処法が入っている（docs/ERRORS.md も参照）
        logger.error("処理を中断しました: %s", e)
        raise
    except Exception:
        logger.error("予期しないエラーが発生しました", exc_info=True)
        raise