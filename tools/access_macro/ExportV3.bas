Attribute VB_Name = "ExportV3"
Option Explicit

' === ここを実行環境に合わせて書き換える ===
'
' Access データベースのフルパス
Public Const DB_PATH As String = "C:\作業\database.accdb"
'
' エクスポート対象のテーブルまたはクエリ名
Public Const SOURCE_NAME As String = "Q_一覧"
'
' 出力先フォルダ（既存の INPUT_FOLDER と一致させる）
Public Const OUTPUT_FOLDER As String = "C:\作業\input"
'
' 出力ファイル名の接頭辞（Access 側由来 "V3_"、既存の "一覧_" などへ切り替える場合は書き換える）
Public Const FILE_PREFIX As String = "V3_"
' ===

Public Sub ExportV3()
    ' 1. 当日日付のファイル名を組み立てる: V3_YYYYMMDD.xlsx
    Dim fileName As String
    fileName = FILE_PREFIX & Format(Date, "yyyymmdd") & ".xlsx"
    Dim outputPath As String
    outputPath = OUTPUT_FOLDER & "\" & fileName

    On Error GoTo Err_ExportV3

    ' 2. Access の TransferSpreadsheet でエクスポート
    '    - acExport = エクスポート
    '    - acSpreadsheetTypeExcel12Xml = .xlsx 形式
    '    - 既存の同名ファイルは上書き
    DoCmd.TransferSpreadsheet _
        TransferType:=acExport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        TableName:=SOURCE_NAME, _
        FileName:=outputPath, _
        HasFieldNames:=True

    ' 3. 完了通知（手動実行時は MsgBox、ログ用に Debug.Print も出す）
    Debug.Print "Exported: " & outputPath
    MsgBox "エクスポート完了: " & outputPath, vbInformation, "ExportV3"

    Exit Sub

Err_ExportV3:
    MsgBox "エクスポートに失敗しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, _
           vbCritical, "ExportV3 Error"
    Debug.Print "ExportV3 failed: " & Err.Number & " - " & Err.Description
    Resume Exit_ExportV3

Exit_ExportV3:
End Sub
