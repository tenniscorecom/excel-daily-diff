@echo off
chcp 65001 >nul
rem ============================================================
rem  config.ini から補完用スタブ（typings/comken/）を作り直す。
rem  config.ini のセクション・キーを増やしたあとに実行すると、
rem  VS Code（Pylance）で config.SECTION.KEY が補完されるようになる。
rem
rem  ※ 普段は main.py を1回動かせば自動生成されるので必須ではない。
rem     「ツールを動かさずに補完だけ先に用意したい」ときに使う。
rem ============================================================

set "COMKEN_ROOT=\\server\share\tools\comken"

cd /d "%~dp0"
set "PYTHONPATH=%COMKEN_ROOT%;%PYTHONPATH%"

python -m comken.config
if errorlevel 1 (
  echo.
  echo [!] スタブ生成に失敗しました。config.ini があるか確認してください。
  pause
)
