Attribute VB_Name = "ResultView"
Option Explicit

Private Const RESULT_TABLE As String = "tblCsv"

' 結果表示フォームを開く。フォームが既に開いているときは、最新の CSV を再読み込みする。
Public Sub ShowResult()
    Dim wasOpen As Boolean

    On Error GoTo Err_ShowResult
    wasOpen = CurrentProject.AllForms("frmResult").IsLoaded
    DoCmd.OpenForm "frmResult", acNormal, "", "", , acNormal
    ' Form_Load は初回表示時に ReloadResult を呼ぶ。既に開いている場合は
    ' Form_Load が走らないため、ここで明示的に再読み込みする。
    If wasOpen Then ReloadResult
    Exit Sub

Err_ShowResult:
    MsgBox "結果フォームを開けませんでした。" & vbCrLf & _
           "frmResult が作成されているか、" & OUTPUT_FOLDER & "\" & OUTPUT_NAME & _
           " が存在するか確認してください。" & vbCrLf & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical, "ShowResult Error"
End Sub

' frmResult.Form_Load から呼ばれる。CSV を固定リンクテーブルへ取り込み、
' frmResult の RecordSource に設定する。
Public Sub ReloadResult()
    Dim csvPath As String
    Dim resultForm As Object

    On Error GoTo Err_ReloadResult
    If Right$(OUTPUT_FOLDER, 1) = "\" Then
        csvPath = OUTPUT_FOLDER & OUTPUT_NAME
    Else
        csvPath = OUTPUT_FOLDER & "\" & OUTPUT_NAME
    End If
    If Len(Dir$(csvPath)) = 0 Then
        Err.Raise 53, "ReloadResult", "CSV ファイルが見つかりません: " & csvPath
    End If

    ' フォームをいったん非連結にしてからリンクを張り直す。
    Set resultForm = Forms("frmResult")
    resultForm.RecordSource = ""
    DropResultLink

    ' DiffPipeline.bas が BOM 付き UTF-8 で出力するため、Access 標準の
    ' 区切りテキストリンクで日本語の見出しをそのまま読み込める。
    DoCmd.TransferText _
        TransferType:=acLinkDelim, _
        SpecificationName:="", _
        TableName:=RESULT_TABLE, _
        FileName:=csvPath, _
        HasFieldNames:=True

    resultForm.RecordSource = RESULT_TABLE
    resultForm.Requery
    Exit Sub

Err_ReloadResult:
    MsgBox "結果の再読み込みに失敗しました。" & vbCrLf & _
           "対象ファイル: " & csvPath & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical, "ReloadResult Error"
End Sub

' 結果 CSV 用のリンクテーブルを削除する。ローカルテーブルを誤って消さないよう、
' 同名のローカルテーブルがある場合はエラーにする。
Private Sub DropResultLink()
    Dim db As Object
    Dim tableDef As Object

    Set db = CurrentDb
    On Error Resume Next
    Set tableDef = db.TableDefs(RESULT_TABLE)
    On Error GoTo 0
    If tableDef Is Nothing Then Exit Sub

    If Len(Nz(tableDef.Connect, "")) = 0 Then
        Err.Raise vbObjectError + 2101, "DropResultLink", _
                  "ローカルテーブル " & RESULT_TABLE & " が既に存在します。" & _
                  "結果表示用のリンクテーブル名を空けてください。"
    End If

    db.TableDefs.Delete RESULT_TABLE
    db.TableDefs.Refresh
End Sub
