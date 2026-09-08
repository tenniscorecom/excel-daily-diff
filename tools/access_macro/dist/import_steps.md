# 起動 UI の組み込み手順

Access DB（`ExportV3.bas` / `DiffPipeline.bas` をインポート済み）に、
**Access を開いただけで起動フォーム `frmStart` が表示される** AutoExec 仕組みを
組み込むまでの手順。

`UILauncher.bas` まで含めて全部入りの `.accdb` 完成版が手元にある場合は
「2. 完成版 .accdb を開く」だけ。1 から組み上げる場合は「3. 手動で組み上げる」。

---

## 1. 前提

- Access 2010 以降（`acSpreadsheetTypeExcel12Xml` が .xlsx 形式用）
- `ExportV3.bas` / `DiffPipeline.bas` の **両方** が同じ DB に
  インポート済み（`DiffPipeline` は `ExportV3` の `Public Const` を参照するので
  片方だけでは動かない）
- DB のマクロ設定が**有効**（Access の `ファイル → オプション → トラストセンター
  → マクロの設定` で「すべてのマクロを有効にする」または「通知してマクロを無効
  にする」のいずれか。AutoExec 起動時に通知バーで許可が必要なら許可する）

---

## 2. 完成版 .accdb を開く

`.accdb` 完成版が手元にある場合:

1. エクスプローラから `enki.accdb` をダブルクリックして Access で開く
2. **マクロの有効化** を求められたら「コンテンツを有効にする」をクリック
3. `frmStart` が自動で表示されれば完了

完了。3 は飛ばす。

---

## 3. 手動で組み上げる

`.accdb` 完成版がなく、`ExportV3.bas` / `DiffPipeline.bas` / `UILauncher.bas` /
`frmStart_spec.md` だけがある場合の手順。

### 3.1 標準モジュールのインポート

1. Access で対象の DB を開く
2. `Alt + F11` で VBE を開く
3. **ファイル → ファイルのインポート** で `ExportV3.bas` を選択
   - プロジェクトに `ExportV3` モジュールが追加される
4. 同じ手順で `DiffPipeline.bas` をインポート
5. 同じく `UILauncher.bas` をインポート（AutoExec を後付けできない環境で手動起動
   したいときのため。**省略可**）
6. `ResultView.bas` をインポート（「結果を表示」ボタンと `frmResult` の
   CSV 読み込みに必要）

> `UILauncher.bas` は AutoExec が既に動く環境では使われない。1 度フォームを
> 閉じてから Immediate Window で `Call LaunchUI` して再オープンする場面用に
> 置いてあるだけ。

### 3.2 起動フォーム `frmStart` の作成

詳細は [`frmStart_spec.md`](./frmStart_spec.md) を参照。概略:

1. VBE ではなく Access 側で **作成 → フォーム デザイン** を選択
2. フォーム名を `frmStart` にする（プロパティシートの「その他」タブ）
3. 標題を `延期積上集計 - 起動` にする（プロパティシートの「書式」タブ）
4. 仕様書の §2 に従い、ラベル 14 個とコマンドボタン 3 個を配置
   - ラベル: `lblTitle` / `lblSection1` / `lblDbPath` / `lblSourceName` /
     `lblOutputFolder` / `lblFilePrefix` / `lblApplyFilter` / `lblTargetMonth` /
     `lblSection2` / `lblRollingMode` / `lblRollingDays` / `lblKeyColumn` /
     `lblIncrementalSave` / `lblHint`
   - ボタン: `btnExport` / `btnRunPipeline` / `btnShowResult`
5. ボタンの **On Click** を **マクロ式** にする:
   - `btnExport` の On Click → `=ExportV3()`
   - `btnRunPipeline` の On Click → `=RunFullPipeline()`
   - `btnShowResult` の On Click → `=ShowResult()`
6. フォームを保存して閉じる

### 3.3 フォームモジュールのコード貼り付け

1. ナビゲーションウィンドウで `frmStart` を右クリック → **コードの表示**
2. [`frmStart_spec.md` §4](./frmStart_spec.md#4-フォームモジュールform_frmstartの中身)
   のコードを貼り付けて保存
3. フォームビュー（F5）で開いて、ラベル 11 個に `Public Const` の現在値が
   表示されることを確認

### 3.4 AutoExec マクロの作成

1. Access の **作成 → マクロ** を選択
2. デザイングリッドで **新しいアクション** を開く → **OpenForm** を選択
3. フォーム名: `frmStart`
4. ビュー: `フォーム`
5. フィルタ名 / Where 条件式: **空**（既定）
6. データモード: **編集**（既定）
7. ウィンドウモード: **標準**（既定）
8. 右クリック → **名前を付けて保存** で **マクロ名 `AutoExec`** にして保存
   - **重要**: マクロ名は必ず **`AutoExec`** にする。Access 起動時に自動実行
     されるのはこの名前のマクロだけ
9. セキュリティ警告バーで「**マクロをすべて有効化**」をクリック
   - 一度有効化すれば次回以降は出ない

### 3.5 動作確認

1. Access を一旦閉じる
2. DB を再度開く
3. ナビゲーションウィンドウの上に `frmStart` が自動で表示されれば OK
4. それぞれのボタンを押して、`ExportV3` / `RunFullPipeline` が走れば完了

### 3.6 結果表示フォームの組み上げ

`ResultView.bas` をインポートした後、[`frmResult_spec.md`](./frmResult_spec.md) に従って
結果表示フォームを作る。

1. Access 側で **作成 → フォーム デザイン**を選択する。
2. フォーム名を `frmResult`、標題を `集計結果` にする。
3. Default View を `Datasheet`、Allow Datasheet View / Allow Form View を `Yes` にする。
4. Scroll Bars を `Both`、Record Selectors を `No`、Navigation Buttons を `Yes`、
   ポップアップを `No` にする。Close Button は有効のままにする。
5. `Record Source` は空欄のままでもよい（`ReloadResult` がリンク後に `tblCsv` を設定する）。
   フォーム上にコントロールを配置する必要はない。
6. フォームのコードに次を貼り付けて保存する。

   ```vba
   Option Compare Database
   Option Explicit

   Private Sub Form_Load()
       ReloadResult
   End Sub
   ```

7. `frmStart` のデザインビューを開き、`btnRunPipeline` の右隣にコマンドボタンを追加する。
   コントロール名を `btnShowResult`、標題を `結果を表示` にし、タブ順を最後にする。
8. `btnShowResult` の On Click にマクロ式 `=ShowResult()` を設定する。
9. `集計.csv` が作成済みの状態で `frmStart` の「結果を表示」を押し、
   `frmResult` がデータシートビューで表示されることを確認する。

初回の「結果を表示」時に `OUTPUT_FOLDER\集計.csv` が `tblCsv` としてリンクされる。
以後もボタンを押すたびにリンクを張り直すため、集計後の最新内容が表示される。
既存のローカルテーブル `tblCsv` がある場合は誤削除を避けて停止するので、先に名前を変更する。

---

## 4. 設定変更（`Public Const`）の流れ

1. `Alt + F11` で VBE を開く
2. `ExportV3` モジュール（または `DiffPipeline`）を開く
3. 該当する `Public Const` を書き換える
4. Access を **再起動**（保存だけでは反映されない。`RefreshLabels` は起動時に
   1 回値を読み込むだけ）
5. 起動フォームのラベル表示が新値に変わっていることを確認

---

## 5. トラブルシューティング

| 症状 | 原因と対処 |
|---|---|
| Access を開いてもフォームが出ない | AutoExec マクロの保存名が違う（必ず `AutoExec`）。またはマクロが無効化されている（セキュリティ警告で許可） |
| ボタンが「パラメータの入力」を求めてくる | On Click が `=ExportV3()` でなく `ExportV3` になっている。等号を忘れない |
| ボタンが押しても無反応 | `ExportV3` / `RunFullPipeline` の `Sub` が Public でなく Private になっていないか確認（既存 .bas は Public のはず） |
| ラベルに `#Name?` と表示される | リフレッシュ前の失敗。`F5` でフォームを再表示。それでも出るなら `Public Const` 名が間違っている |
| 起動時 `frmStart` が見つからないエラー | AutoExec の OpenForm が `frmStart` を指定しているがフォーム未作成。先に 3.2 / 3.3 を完了させる |
| `frmStart` を閉じた後 VBA を再実行したい | Immediate Window で `Call LaunchUI`（`UILauncher.bas` インポート済みの場合）または VBE で `Sub LaunchUI` にカーソル → F5 |
