# Access 用 VBA マクロ

Microsoft Access から当日付のデータを Excel にエクスポートする VBA マクロです。
エクスポート先は Python 側 (`src/run.py`) が読み取る `INPUT_FOLDER` を想定しています。

書き出したファイルを Access 内で読み直して集計まで済ませたいときは、
`DiffPipeline.bas`（後述）を使うと Python 無しで `集計.csv` まで作れます。

## ファイル

| ファイル | 役割 |
|---|---|
| `ExportV3.bas` | Access 側で動く VBA モジュール本体（Excel への書き出し） |
| `DiffPipeline.bas` | 書き出し済みファイルを読み直して集計まで行う（Python 版と同じ処理） |

## 概要

- Access DB（テーブル / クエリ）の中身を `YYYYMMDD.xlsx` 形式で書き出す
- 既存の同名ファイルは上書きされる（差分取り込み運用のため、毎回同じファイル名を使う）
- 完了時は MsgBox で通知し、Immediate Window にもログを出す

## インポート手順

1. Access で対象の DB を開く
2. `Alt + F11` で Visual Basic Editor（VBE）を開く
3. **ファイル → ファイルのインポート** で `ExportV3.bas` を選択
4. モジュール `ExportV3` がプロジェクトに追加される
5. 集計まで Access で行う場合は、同じ手順で `DiffPipeline.bas` も追加する
   （`DiffPipeline` は `ExportV3` の定数を参照するので、片方だけでは動かない）

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

---

# DiffPipeline.bas（Access 単体で集計まで終わらせる）

Python ツール（`src/run.py`）と同じ「延期・積み上げ集計」を Access VBA 内で完結させる
パイプライン。**Python が入っていない PC でも Access 単体で `集計.csv` を作れる**。
`ExportV3.bas` が「Access → Excel の書き出し」までを担うのに対し、こちらは**書き出し
済みの `V3_YYYYMMDD.xlsx` を読み直して集計する**。両方を同じ DB に入れておけば、
エクスポートから集計まで Access だけで回る。

## 実行方法

VBE で `Sub RunFullPipeline` 内にカーソルを置いて `F5`、または Access の イミディエイト
ウィンドウで `Call RunFullPipeline`。進捗と業務日ごとの件数は `Debug.Print` で
イミディエイト ウィンドウに出て、完了時・失敗時は MsgBox で通知する。

## 前提

- `ExportV3.bas` と `DiffPipeline.bas` の**両方**を同じ DB にインポートしておく
  （`DiffPipeline` は `ExportV3` の `Public Const` を参照するので片方だけでは動かない）
- `OUTPUT_FOLDER`（= Python の `[FILES] INPUT_FOLDER`）に `V3_YYYYMMDD.xlsx` が日付ぶん
  並んでいて、`FILE_PREFIX` がその接頭辞と一致している（ここでは**読み取り元**）
- 参照設定の追加は不要（`Scripting.Dictionary` / `FileSystemObject` / `VBScript.RegExp` /
  `ADODB.Stream` はすべて遅延バインディングで作る）

## 定数の書き換え

`ExportV3.bas` の 11 定数（上段 4 + フィルタ用 7）はそのまま使う。絞り込みは
`APPLY_FILTER` の値によらず**常に適用される**（集計の前提条件のため）。
`DiffPipeline.bas` 側で書き換えるのは次の 7 つ（`DATE_PATTERN` は通常触らない）:

| 定数 | 意味 | 既定値 |
|---|---|---|
| `KEY_COLUMN` | 前日と当日を突き合わせるキー列（Python の `[SOURCE] KEY_COLUMN`） | `顧客番号` |
| `ROLLING_MODE` | `False`=年次累積モード / `True`=ローリングモード | `False` |
| `ROLLING_WINDOW_DAYS` | ローリングモードの窓幅（日数） | `7` |
| `INCREMENTAL_SAVE` | 1日ぶん終わるたびに CSV を途中保存するか | `True` |
| `CLEANUP_TEMP_TABLES` | 一時テーブルを最後に必ず削除するか | `True` |
| `OUTPUT_NAME` | 出力 CSV のファイル名 | `集計.csv` |
| `TEMP_TABLE_PREFIX` | 一時テーブル名の接頭辞 | `_V3_` |

## 実行モード

| モード | 設定 | 読む範囲（業務日） |
|---|---|---|
| 年次累積 | `ROLLING_MODE = False`（既定） | 実行年の 1/1 〜 今日 |
| ローリング | `ROLLING_MODE = True` | 今日 -(`ROLLING_WINDOW_DAYS` - 1)日 〜 今日 |

年次累積モードでは対象月が業務日の月だけなので、`来月` 行は構造上できるが中身は空のまま。
ローリングモードでは業務日の当月と翌月の両方を対象月にする（12 月の翌月は翌年 1 月）。

## 処理の流れ

1. `OUTPUT_FOLDER` の `FILE_PREFIX*.xlsx` を集め、ファイル名の日付順に並べる
2. 集計範囲（業務日の下限・上限）を決める
3. 範囲内のファイル + 比較相手として直前 1 本（範囲外）を選ぶ
4. 1 本ずつ `DoCmd.TransferSpreadsheet` で一時テーブル `_V3_YYYYMMDD` へ取り込む
5. 隣り合う 2 本を SQL の `LEFT JOIN` で突き合わせ、`延期` / `積み上げ` を数える
6. `INCREMENTAL_SAVE = True` なら 1 日ぶん終わるたびに CSV を書き出す（`False` なら最後に 1 回）
7. 一時テーブルを削除する（途中で失敗して止まったときも削除する）

**業務日はファイル名の日付 -1日**。入力ファイルは「前日終了時点」のデータなので、
`V3_20260825.xlsx` の中身は 8/24 終了時点であり、`V3_20260824.xlsx` との差分は「8/24 に
動いたぶん」になる（Python 側と同じ）。

## 出力

`OUTPUT_FOLDER` に `OUTPUT_NAME`（既定 `集計.csv`）で書き出す。UTF-8（BOM なし）・CRLF。

- 列は `対象月, 種別, 判定, <業務日...>` の 3 キー列 + 業務日の列
- 行は `当月`/`来月` × `PLAN_PREFIXES` の種別 × `積み上げ`/`延期` の固定構成
  （`PLAN_PREFIXES = 標準,上位` のときは 8 行）
- セルは 3 通り。件数 / `0`（比較したが 0 件）/ 空（比較できていない業務日、またはその
  業務日ではそのラベルが対象外）

## 制約・既知の制限

- **増分実行はしない**。実行するたびに範囲全体を読み直して `集計.csv` を作り直す
  （Python 側の「既存 CSV の最終列から続ける」動作は未対応）
- 範囲内の最初のファイルに比較相手（範囲外の直前 1 本）が無いときは、何も出力せずに
  終わる（Python 側と同じ挙動）。240 件超のファイルでは初回実行に時間がかかる
- 出力先は入力フォルダと同じ `OUTPUT_FOLDER`。Python 側は `[REPORT] OUTPUT_FOLDER` に
  分けているので、置き場を分けたいときは `OutputPath()` を書き換える
- `集計.csv` を Excel で開いたまま実行すると、書き出しに失敗する
- 取り込み先の列名は Excel の見出しそのままになる。見出しに余分な空白が入っていると
  `[顧客番号]` などの列が見つからずエラーになる（Python 側は空白を落とすので違う）
- `DATE_COLUMN` が文字列として取り込まれた場合、`Format([予定日],'yyyy-mm')` が日付として
  整形されず対象月に当たらないことがある（`ExportV3.bas` のフィルタモードと同じ制限）
- `KEY_COLUMN` がファイル内で重複していると、その分だけ件数が多く出る
  （Python 側は最初の 1 行だけ採用して警告を出す）
- `CLEANUP_TEMP_TABLES = False` のときは一時テーブル `_V3_*` が DB に残るので手動で
  削除する（残っていても次回実行時に作り直すので支障はない）
- `[EXCLUDE_CONTAINS]` などの追加絞り込みとシート名の候補指定は未対応（Python 側を使う）
