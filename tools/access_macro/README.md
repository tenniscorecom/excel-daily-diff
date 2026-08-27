# Access エクスポート用 VBA マクロ

Microsoft Access から当日付のデータを Excel にエクスポートする VBA マクロです。
エクスポート先は Python 側 (`src/run.py`) が読み取る `INPUT_FOLDER` を想定しています。

## ファイル

| ファイル | 役割 |
|---|---|
| `ExportV3.bas` | Access 側で動く VBA モジュール本体 |

## 概要

- Access DB（テーブル / クエリ）の中身を `YYYYMMDD.xlsx` 形式で書き出す
- 既存の同名ファイルは上書きされる（差分取り込み運用のため、毎回同じファイル名を使う）
- 完了時は MsgBox で通知し、Immediate Window にもログを出す

## インポート手順

1. Access で対象の DB を開く
2. `Alt + F11` で Visual Basic Editor（VBE）を開く
3. **ファイル → ファイルのインポート** で `ExportV3.bas` を選択
4. モジュール `ExportV3` がプロジェクトに追加される

## 定数の書き換え

`ExportV3.bas` の冒頭にある 4 つの `Public Const` を、実行環境に合わせて書き換えてから使う:

| 定数 | 意味 | 例 |
|---|---|---|
| `DB_PATH` | Access データベースのフルパス | `C:\作業\database.accdb` |
| `SOURCE_NAME` | エクスポート対象のテーブル / クエリ名 | `Q_一覧` |
| `OUTPUT_FOLDER` | 出力先フォルダ（Python 側の `[FILES] INPUT_FOLDER` と一致させる） | `C:\作業\input` |
| `FILE_PREFIX` | 出力ファイル名の接頭辞（V3_ / 一覧_ など） | `V3_` |

ファイル名は `<FILE_PREFIX>YYYYMMDD.xlsx` の形で生成される（業務日 -1日補正は
行わないので、Access 側でそのまま「当日終了時点」を出力する想定）。

## 実行方法

| 方法 | 操作 |
|---|---|
| 手動（VBE） | VBE で `Sub ExportV3` 内にカーソルを置いて `F5` |
| コマンドライン | Access の Immediate Window で `Call ExportV3` |

エラー時は MsgBox でエラー番号と詳細を表示し、`Debug.Print` にも同じ内容が出る。

## 前提

- **Access 2010 以降**（`acSpreadsheetTypeExcel12Xml` は .xlsx 形式用。古い Access では
  別の `SpreadsheetType` 定数が必要）
- Access DB が開ける状態で実行する（`DB_PATH` から自動で開かれるわけではない）
- 出力先フォルダ（`OUTPUT_FOLDER`）が存在し、書き込み権限があること

## Python 側との接続

このマクロは Python プロジェクトの `config.ini` に項目を追加しない。
エクスポート先のパスは VBA の `OUTPUT_FOLDER` 定数のみで管理する。

Python 側は `[FILES] INPUT_FOLDER` を同じパスに合わせ、`FILE_PATTERN` を
`<FILE_PREFIX>*.xlsx` の形に設定すれば、`src/run.py` がそのままこの VBA の
出力ファイルを読み取って集計する（[FILTER] セクションで列名や絞り込み語を
合わせるのが別途必要）。
