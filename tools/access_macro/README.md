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

`ExportV3.bas` 冒頭（上段）の 4 つの `Public Const` を、実行環境に合わせて書き換えてから使う:

| 定数 | 意味 | 既定値（例） |
|---|---|---|
| `DB_PATH` | Access データベースのフルパス | `C:\作業\database.accdb` |
| `SOURCE_NAME` | エクスポート対象のテーブル / クエリ名 | `Q_一覧` |
| `OUTPUT_FOLDER` | 出力先フォルダ（Python 側の `[FILES] INPUT_FOLDER` と一致させる） | `C:\作業\input` |
| `FILE_PREFIX` | 出力ファイル名の接頭辞（V3_ / 一覧_ など） | `V3_` |

ファイル名は `<FILE_PREFIX>YYYYMMDD.xlsx` の形で生成される（業務日 -1日補正は
行わないので、Access 側でそのまま「当日終了時点」を出力する想定）。

### フィルタを使うときの下段 7 定数

`APPLY_FILTER = True` にするときだけ、下段の 7 定数を絞り込み条件に合わせて
書き換える。`False`（既定）のときは無視されるので、そのままで構わない:

| 定数 | 意味 | 既定値 |
|---|---|---|
| `APPLY_FILTER` | Access 側でフィルタを適用するか | `False`（生データ出力） |
| `PLAN_COLUMN` | 種別の列名 | `種別` |
| `PLAN_PREFIXES` | 残す種別の接頭辞群（カンマ区切り） | `標準,上位` |
| `KIND_COLUMN` | 状態の列名 | `状態` |
| `KIND_VALUES` | 残す状態群（カンマ区切り） | `完了,予定` |
| `DATE_COLUMN` | 対象月判定に使う日付列名 | `予定日` |
| `TARGET_MONTH` | 対象月（`yyyy-mm`）。空なら実行日ベース | `` |

## 実行方法

| 方法 | 操作 |
|---|---|
| 手動（VBE） | VBE で `Sub ExportV3` 内にカーソルを置いて `F5` |
| コマンドライン | Access の Immediate Window で `Call ExportV3` |

エラー時は MsgBox でエラー番号と詳細を表示し、`Debug.Print` にも同じ内容が出る。

### フィルタモード（`APPLY_FILTER = True`）

Access 側で Python の `[FILTER]` セクションと同じ絞り込み条件を適用してから
エクスポートするモード。
- 有効にする: `APPLY_FILTER = True` にしてから実行。
- 効果: WHERE 句に `[PLAN_COLUMN] LIKE '標準' OR [PLAN_COLUMN] LIKE '上位'`、
  `[KIND_COLUMN]='完了' OR [KIND_COLUMN]='予定'`、
  `Format([DATE_COLUMN],'yyyy-mm')='2026-08'` を AND で結合する。
- 用途: Python 側の `[FILTER]` を空にして Access に寄せたいときや、Access から
  直接集計用ファイルを作るときに使う。
- 副作用: Access DB に一時クエリ `_TempExportQuery` が作られ、エクスポート後に
  自動削除される（エラー時も `Err_ExportV3:` でクリーンアップする）。

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
