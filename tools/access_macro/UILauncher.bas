Attribute VB_Name = "UILauncher"
Option Explicit

' === 起動 UI の薄いラッパー ===
'
' 既存の ExportV3.bas / DiffPipeline.bas は 1 バイトも触らず、
' Access を起動するだけで AutoExec が起動フォーム frmStart を開く運用が
' 正攻法。.accdb 完成版が手元にないとき（VBA ソース＋仕様書だけで運用
' するとき）は、このモジュールをインポートしてから VBE で
'     Sub LaunchUI にカーソル → F5
' または Immediate Window で
'     Call LaunchUI
' を実行すると、AutoExec の代わりに手動で frmStart を開ける。
'
' フォームを閉じてもこのモジュールは残るため、集計のたびに
'     Call LaunchUI
' を再実行すればよい（DoCmd.OpenForm は同名フォームが既に開いていれば
' それを前面に出すだけ）。

Public Sub LaunchUI()
    ' 起動フォーム frmStart をフォームビューで開く。
    ' AutoExec マクロが有効なら Access 起動時に既に開いているので、
    ' このプロシージャは「AutoExec を後から動かしたい」「.accdb に
    ' AutoExec を仕込む前の手動起動」用途。
    On Error GoTo Err_LaunchUI
    DoCmd.OpenForm "frmStart", acNormal, "", "", , acNormal
    Exit Sub

Err_LaunchUI:
    MsgBox "起動フォームを開けませんでした。" & vbCrLf & _
           "先に frmStart を作成してください（dist/frmStart_spec.md 参照）。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, _
           vbCritical, "LaunchUI Error"
End Sub
