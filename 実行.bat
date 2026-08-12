@echo off
chcp 65001 >nul
rem ============================================================
rem  このツールの起動用。ダブルクリックで main.py を実行します。
rem
rem  初回だけ: 下の COMKEN_ROOT を、共有サーバー上の comken の場所に書き換える。
rem  （PC の環境変数は変更しません。この bat の実行中だけ PYTHONPATH を設定します）
rem ============================================================

set "COMKEN_ROOT=\\server\share\tools\comken"

cd /d "%~dp0"
set "PYTHONPATH=%COMKEN_ROOT%;%PYTHONPATH%"

python main.py

if errorlevel 1 (
  echo.
  echo [!] エラーで終了しました。上の赤い文字（エラー名）を docs\ERRORS.md で調べてください。
  pause
)
