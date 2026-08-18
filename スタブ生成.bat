@echo off
setlocal
rem config.ini から補完用スタブ（typings/comken/）を作り直す。
rem config.ini のセクション・キーを増やしたあとに実行すると、
rem VS Code（Pylance）で config.SECTION.KEY が補完されるようになる。
rem
rem ※ 普段は main.py を1回動かせば自動生成されるので必須ではない。
rem    「ツールを動かさずに補完だけ先に用意したい」ときに使う。

rem comken の場所。PC に恒久登録していない場合だけ、ここが使われる
set "PYTHON_LIBRARY=\\server\share\tools"

rem 共有フォルダ（\\サーバー名\...）から起動されても動くよう pushd を使う（cd は UNC 不可）
pushd "%~dp0" || (
  echo [エラー] このフォルダを開けませんでした: %~dp0
  exit /b 1
)

where python >nul 2>&1 || (
  echo [エラー] Python が見つかりません。
  echo   このパソコンに Python が入っているか、管理者に確認してください。
  popd
  exit /b 1
)

rem すでに PYTHONPATH が通っていれば、そのまま動かす（恒久登録してある場合）
python -c "import comken" >nul 2>&1
if not errorlevel 1 goto :run

rem 通っていないので、この bat に書いてある場所を使う
set "PYTHONPATH=%PYTHON_LIBRARY%;%PYTHONPATH%"

rem 一番多い失敗を先に名指しで出す
if not exist "%PYTHON_LIBRARY%\comken\__init__.py" (
  echo [エラー] 共通ライブラリ comken が見つかりません。
  echo     さがした場所: %PYTHON_LIBRARY%
  echo.
  echo   - 共有サーバーにつながっているか確認してください
  echo   - つながっているなら、この bat の PYTHON_LIBRARY が正しいか確認してください
  popd
  exit /b 1
)

:run
python -m comken config
rem 終了コードは popd より前に控える（popd が成功すると 0 で上書きされる）
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [失敗] スタブ生成を中断しました（終了コード %EXIT_CODE%）。
  echo   エラーの内容は画面の上のほうに出ています。
)

rem 終了コードをそのまま返す
endlocal & exit /b %EXIT_CODE%
