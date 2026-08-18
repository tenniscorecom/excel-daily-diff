@echo off
setlocal
rem このツールの起動用。ターミナルから `実行.bat` を叩くか、`python main.py` を直接実行してください。
rem **RPA 基盤から呼び出すときの入口でもある。** 終了コードをそのまま返すので、
rem スケジューラや RPA が成否を判断できる（pause を入れていないのは、無人実行で止まらないようにするため）。
rem comken を別の場所へ移したときは、ここと .vscode\settings.json の両方を直してください。

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

rem 一番多い失敗を先に名指しで出す。ここで止めないと、後から出る Python のエラーが
rem 「ModuleNotFoundError: comken」だけになり、原因が共有サーバーだと分からない
if not exist "%PYTHON_LIBRARY%\comken\__init__.py" (
  echo [エラー] 共通ライブラリ comken が見つかりません。
  echo     さがした場所: %PYTHON_LIBRARY%
  echo.
  echo   - 共有サーバーにつながっているか確認してください
  echo   - つながっているなら、この bat の PYTHON_LIBRARY が正しいか確認してください
  echo   - このパソコンで何度も使うなら、comken のフォルダにある
  echo     setup_comken.bat を1回実行しておくと、以後この bat を直さずに済みます
  popd
  exit /b 1
)

:run
python main.py
rem 終了コードは popd より前に控える（popd が成功すると 0 で上書きされる）
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [失敗] 処理を中断しました（終了コード %EXIT_CODE%）。
  echo   エラーの内容は画面の上のほうに出ています。logs フォルダにも残っています。
  echo   分からないときは、この画面をスクリーンショットして管理者へ送ってください。
)

rem 終了コードをそのまま返す。スケジューラや RPA 基盤が成否を判断できるようにする
endlocal & exit /b %EXIT_CODE%
