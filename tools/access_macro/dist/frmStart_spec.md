# 起動フォーム `frmStart` 仕様書

`ExportV3.bas` / `DiffPipeline.bas` をインポート済みの Access DB に、
**Access を起動するだけで AutoExec が起動フォームを開く** 仕組みを被せるための
フォーム仕様。`Public Const` はコンパイル時定数のため実行時に書き換えできない
ので、入力欄は持たず、**現在値をラベル表示するだけ**にする。設定変更は引き続き
VBE での `Public Const` 編集（初回 1 回だけ）。

ボタン 3 つ（エクスポート実行 / 集計実行 / 結果を表示）が On Click で既存関数を直接呼ぶ。
既存のエクスポート / 集計ボタン用のラッパー VBA は設けず、結果表示に必要な
`ResultView.bas` の `ShowResult` だけを追加する。

---

## 1. フォーム プロパティ

| プロパティ | 値 | 備考 |
|---|---|---|
| フォーム名 | `frmStart` | コードから `DoCmd.OpenForm` するので必須 |
| 標題 | `延期積上集計 - 起動` | タイトルバー表示 |
| Default View | `Single Form` | 単票フォーム |
| Allow Datasheet View | `No` | 帳票ビューには切替できない |
| Allow Form View | `Yes` | フォームビューは許可 |
| Scroll Bars | `Neither` | 装飾を切る |
| Record Selectors | `No` | 装飾を切る |
| Navigation Buttons | `No` | 装飾を切る |
| Close Button | 有効 | × ボタンで閉じられるようにする |
| ControlBox | 有効 | 最小化 / 最大化も有効（既定のまま） |
| ポップアップ | `No` | 通常の子ウィンドウとして表示 |
| 作業ウィンドウ固定 | `No` | 後ろのウィンドウも触れるようにする |
| 自動中央寄せ | `いいえ` | 任意（既定のまま） |

> ナビゲーションウィンドウの表示は変更しない。フォームを閉じた後でも VBA を
> 再実行できるようにするため、`StartupShowDBWindow = False` 等の設定は使用しない。

---

## 2. コントロール一覧

Access のフォームデザイナで **デザインビュー**を開き、ツールボックスから
順番に配置する。`lblSection*` などの見出しラベルはフォントを太字にすると
見栄えが良い（必須ではない）。

| # | コントロール名 | 型 | ラベル / 標題 | 値（既定） | 位置の目安 | 紐づく対象 |
|---|---|---|---|---|---|---|
| 1 | `lblTitle` | ラベル | `延期積上集計` | —（フォント大きめ、太字） | 上端中央 | — |
| 2 | `lblSection1` | ラベル | `▼ エクスポート設定（ExportV3.bas）` | — | 1 の下 | — |
| 3 | `lblDbPath` | ラベル | `Access DB パス: ` | `DB_PATH` の現在値（`Form_Load` で `RefreshLabels` が設定） | 左寄せ、幅広 | `ExportV3.DB_PATH` |
| 4 | `lblSourceName` | ラベル | `テーブル/クエリ名: ` | `SOURCE_NAME` の現在値 | 同上 | `ExportV3.SOURCE_NAME` |
| 5 | `lblOutputFolder` | ラベル | `出力先フォルダ: ` | `OUTPUT_FOLDER` の現在値 | 同上 | `ExportV3.OUTPUT_FOLDER` |
| 6 | `lblFilePrefix` | ラベル | `ファイル名接頭辞: ` | `FILE_PREFIX` の現在値 | 同上 | `ExportV3.FILE_PREFIX` |
| 7 | `lblApplyFilter` | ラベル | `フィルタ適用: ` | `APPLY_FILTER` の現在値 | 同上 | `ExportV3.APPLY_FILTER` |
| 8 | `lblTargetMonth` | ラベル | `対象月: ` | `TARGET_MONTH` の現在値（空なら「当月自動」） | 同上 | `ExportV3.TARGET_MONTH` |
| 9 | `lblSection2` | ラベル | `▼ 集計設定（DiffPipeline.bas）` | — | 8 の下 | — |
| 10 | `lblRollingMode` | ラベル | `ローリングモード: ` | `ROLLING_MODE` の現在値 | 同上 | `DiffPipeline.ROLLING_MODE` |
| 11 | `lblRollingDays` | ラベル | `ローリング窓幅（日）: ` | `ROLLING_WINDOW_DAYS` の現在値 | 同上 | `DiffPipeline.ROLLING_WINDOW_DAYS` |
| 12 | `lblKeyColumn` | ラベル | `突合キー列: ` | `KEY_COLUMN` の現在値 | 同上 | `DiffPipeline.KEY_COLUMN` |
| 13 | `lblIncrementalSave` | ラベル | `1日ごと途中保存: ` | `INCREMENTAL_SAVE` の現在値 | 同上 | `DiffPipeline.INCREMENTAL_SAVE` |
| 14 | `lblHint` | ラベル | `※設定変更は VBE で Public Const を書き換えてから Access を再起動してください。` | — | 下端 | — |
| 15 | `btnExport` | コマンドボタン | `エクスポート実行` | —（On Click: `=ExportV3()`） | 右下寄り、幅広 | `ExportV3.ExportV3` |
| 16 | `btnRunPipeline` | コマンドボタン | `集計実行` | —（On Click: `=RunFullPipeline()`） | `btnExport` の右隣 | `DiffPipeline.RunFullPipeline` |
| 17 | `btnShowResult` | コマンドボタン | `結果を表示` | —（On Click: `=ShowResult()`） | `btnRunPipeline` の右隣 | `ResultView.ShowResult` |

### タブ順

1 → 2 → 3 → ... → 14 → 15 → 16 → 17 の順（上から下、左から右）。

---

## 3. ボタンの On Click イベント

**イベントプロシージャではなく、マクロ式を埋め込む。** Access のボタン On Click
ダイアログで `[イベント プロシージャ]` ではなく `[マクロ ビルダ]` / テキスト直接
入力を選び、以下の式を貼り付ける。

| ボタン | On Click の式 |
|---|---|
| `btnExport` | `=ExportV3()` |
| `btnRunPipeline` | `=RunFullPipeline()` |
| `btnShowResult` | `=ShowResult()` |

> **判断**: マクロ式が動かない環境（？）を心配する声があるが、Access のボタン
> On Click で式が動かない環境は基本的に無い。動かない環境に当たった場合のみ
> イベントプロシージャに切り替える:
>
> ```vba
> Private Sub btnExport_Click()
>     ExportV3
> End Sub
>
> Private Sub btnRunPipeline_Click()
>     RunFullPipeline
> End Sub
>
> Private Sub btnShowResult_Click()
>     ShowResult
> End Sub
> ```

---

## 4. フォームモジュール（`Form_frmStart`）の中身

VBE のプロジェクトウィンドウでフォーム `frmStart` を右クリック →
**コードの表示** を選び、以下のコードを貼り付ける。**Public Const への
書き込みは行わない**（読み取りだけ）。

```vba
Option Compare Database
Option Explicit

' Access 起動時に AutoExec から開かれた直後、または DoCmd.OpenForm で
' 開かれた直後に走る。Public Const の現在値を読み出してラベルに反映する。
Private Sub Form_Load()
    RefreshLabels
End Sub

' フォームの再表示（F5 など）でも値を再読み込みできるように、
' Form_Current も RefreshLabels を呼ぶ。
Private Sub Form_Current()
    RefreshLabels
End Sub

' Public Const の値を読み取ってラベルに設定する。書き込みは一切しない。
' 設定変更は VBE 側で Public Const を編集 → Access 再起動で反映する運用。
Private Sub RefreshLabels()
    Me.lblDbPath.Caption = "Access DB パス: " & DB_PATH
    Me.lblSourceName.Caption = "テーブル/クエリ名: " & SOURCE_NAME
    Me.lblOutputFolder.Caption = "出力先フォルダ: " & OUTPUT_FOLDER
    Me.lblFilePrefix.Caption = "ファイル名接頭辞: " & FILE_PREFIX
    Me.lblApplyFilter.Caption = "フィルタ適用: " & APPLY_FILTER
    Me.lblTargetMonth.Caption = "対象月: " & _
        IIf(Len(TARGET_MONTH) > 0, TARGET_MONTH, "当月自動")
    Me.lblRollingMode.Caption = "ローリングモード: " & ROLLING_MODE
    Me.lblRollingDays.Caption = "ローリング窓幅（日）: " & ROLLING_WINDOW_DAYS
    Me.lblKeyColumn.Caption = "突合キー列: " & KEY_COLUMN
    Me.lblIncrementalSave.Caption = "1日ごと途中保存: " & INCREMENTAL_SAVE
End Sub
```

> **実装方針の判断**: `RefreshLabels` は **Public Const を直接読む**。VBA
> の `Public Const` は Access の式（`=OUTPUT_FOLDER` 等）からは参照
> できないため、ラベルの `Control Source` 経由では値を取れない。VBA 側で
> 読むしかないので `RefreshLabels` 方式を採用した。別関数で取得する
> ラッパーを増やすと、Public Const を「読み取り専用 API」化する改変に
> なってしまうので避ける。

---

## 5. 完了条件（Access での確認）

- [ ] Access を開くと、ナビゲーションウィンドウの上に `frmStart` が表示される
- [ ] ラベル 11 個（lblDbPath / lblSourceName / ... / lblIncrementalSave）に
      既存の `Public Const` の現在値が表示される
- [ ] 「エクスポート実行」ボタンを押すと `ExportV3` が走る（MsgBox で完了通知）
- [ ] 「集計実行」ボタンを押すと `RunFullPipeline` が走る（MsgBox で完了通知）
- [ ] 「結果を表示」ボタンを押すと `frmResult` が開き、最新の `集計.csv` が
      データシートビューで表示される
- [ ] フォームを × で閉じた後、再度 Access を開くと AutoExec がフォームを開く
- [ ] VBE で `Public Const` を書き換えて Access を再起動すると、ラベルの表示も
      新しい値に変わっている

---

## 6. Access で .accdb を一から作る場合の手順

`UILauncher.bas` 等の .bas は手元にあるが、`.accdb` を初めて作る場合は次の
手順で組み上げる。既に .accdb がある場合は [import_steps.md](./import_steps.md)
に進むこと。

1. Access を起動し、空のデータベースを作成（例: `enki.accdb`）
2. `Alt + F11` で VBE を開く
3. 既存の `ExportV3.bas` / `DiffPipeline.bas` をインポートする
   （`ファイル → ファイルのインポート`）。`btnShowResult` の On Click で使う
   `ResultView.bas` も同時にインポートする
   結果表示を使う場合は、同じ手順で `ResultView.bas` もインポートする
4. 本仕様書 §1〜§3 に従ってフォーム `frmStart` を作成する
5. §4 のコードをフォームモジュールに貼り付ける
6. ナビゲーションウィンドウで `frmStart` を右クリック → デザインビューで開き、
   一度フォームビュー（F5）で起動確認
7. [import_steps.md §3「AutoExec マクロの作成」](./import_steps.md#3-autoexec-マクロの作成)
   に従って AutoExec マクロを作成する
8. Access を再起動して、フォームが自動表示されることを確認する
