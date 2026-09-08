# 結果表示フォーム `frmResult` 仕様書

`ExportV3.bas` / `DiffPipeline.bas` をインポートした Access DB に、
`OUTPUT_FOLDER\集計.csv` の内容をデータシートビューで表示するフォームを追加する。
`.accdb` はリポジトリに含めず、この仕様書と `ResultView.bas` を使って Access 側で組み上げる。

---

## 1. フォーム プロパティ

| プロパティ | 値 | 備考 |
|---|---|---|
| フォーム名 | `frmResult` | `ShowResult` / `ReloadResult` がこの名前を参照する |
| 標題 | `集計結果` | タイトルバー表示 |
| Default View | `Datasheet` | CSV と同じ表形式で表示する |
| Allow Datasheet View | `Yes` | データシートビューを許可 |
| Allow Form View | `Yes` | 既定のフォームビューも許可 |
| Record Source | `tblCsv` | `Form_Load` の `ReloadResult` がリンク後に設定する（設計時は空欄でもよい） |
| Scroll Bars | `Both` | 縦横両方 |
| Record Selectors | `No` | 行セレクターを表示しない |
| Navigation Buttons | `Yes` | レコード移動を許可 |
| Close Button | 有効 | × ボタンで閉じられるようにする |
| ポップアップ | `No` | 通常の子ウィンドウとして表示 |

フォーム上にコントロールを手作業で配置する必要はない。CSV の見出しから
Access がデータシートの列を生成する。

---

## 2. フォームの作成

1. Access 側で **作成 → フォーム デザイン**を選択する。
2. フォームのプロパティシートで、上表のフォーム名・標題・表示プロパティを設定する。
3. `Record Source` は空欄のまま保存してもよい。`tblCsv` がまだ存在しない場合は、
   `ReloadResult` が初回表示時に作成してから設定する。
4. フォームを `frmResult` という名前で保存する。
5. `frmResult` を右クリックして **コードの表示**を選び、次のフォームモジュールを貼り付ける。

```vba
Option Compare Database
Option Explicit

' フォームを開くたびに OUTPUT_FOLDER\OUTPUT_NAME の最新内容を表示する。
Private Sub Form_Load()
    ReloadResult
End Sub
```

`ReloadResult` は標準モジュール `ResultView` の Public Sub である。
フォームモジュール側で CSV パスやテーブルを直接指定しない。

---

## 3. CSV の取り込み方法（選択肢 C）

この実装では、固定名のリンクテーブル **`tblCsv`** を使い、クエリは作成しない。
`ResultView.ReloadResult` の処理は次のとおり。

1. `ExportV3.OUTPUT_FOLDER` と `DiffPipeline.OUTPUT_NAME` を読み取り、
   `OUTPUT_FOLDER\OUTPUT_NAME` のフルパスを作る。定数の書き換えは行わない。
2. CSV の存在を確認する。
3. `frmResult` をいったん非連結にする。
4. 既存の `tblCsv` がリンクテーブルなら削除する。ローカルテーブルが同名で存在する
   場合は、誤削除を避けるためエラーにする。
5. `DoCmd.TransferText` の `acLinkDelim` で CSV を `tblCsv` としてリンクする
   （`HasFieldNames:=True`）。
6. `frmResult.RecordSource = "tblCsv"` として `Requery` する。

`DiffPipeline.bas` は UTF-8 BOM 付きの CSV を出力するため、Access 標準の区切り
テキストリンクで日本語の見出しを読み込める。CSV を更新した後に「結果を表示」を
押すたび、リンクを張り直して最新の列構成・値を読み込む。

`tblCsv` は結果表示専用の名前である。既存のローカルテーブルをこの名前で使っている
場合は、フォームを使う前に名前を変更するか、結果表示用に別の名前を空ける。

---

## 4. 列構成

`DiffPipeline.bas` が出力する CSV の見出しをそのままデータシートに表示する。

| 列 | 内容 |
|---|---|
| 1 | `対象月`（`当月` / `来月`） |
| 2 | `種別`（`PLAN_PREFIXES` の接頭辞） |
| 3 | `判定`（`積み上げ` / `延期`） |
| 4 以降 | 業務日（`yyyy-mm-dd`）ごとの件数 |

行数と日付列数は `DiffPipeline` の設定・入力ファイルにより変わる。
結果フォーム側で列を固定定義しないため、CSV の横軸をそのまま確認できる。

---

## 5. 呼び出し元と操作順

- `frmStart.btnShowResult` の **On Click**: `=ShowResult()`
- `ResultView.ShowResult`: `DoCmd.OpenForm "frmResult"` を呼ぶ。
  フォームが既に開いている場合は、開き直さず `ReloadResult` を明示的に呼ぶ。
- `frmResult.Form_Load`: `ReloadResult` を呼び、初回表示時に最新 CSV を反映する。

通常は次の順に操作する。

1. `frmStart` の「集計実行」を押して `集計.csv` を更新する。
2. 「結果を表示」を押す。
3. `frmResult` のデータシートビューで集計結果を確認する。

CSV がまだ存在しない場合や、`tblCsv` を作成できない場合は、`ReloadResult Error`
のメッセージボックスにエラー番号と詳細が表示される。

---

## 6. Access 側での確認項目

- [ ] フォーム名が `frmResult` になっている
- [ ] 標題が `集計結果` になっている
- [ ] Default View が `Datasheet` になっている
- [ ] `Form_Load` が `ReloadResult` を呼んでいる
- [ ] `frmStart` の `btnShowResult` の On Click が `=ShowResult()` になっている
- [ ] `集計.csv` が存在する状態でフォームを開くと、`tblCsv` が作成されて表が表示される
- [ ] CSV を更新して再度「結果を表示」を押すと、最新の内容に置き換わる

Access 実機での確認は会社 PC で行う前提であり、このリポジトリでは静的な VBA と
仕様書を納品する。
