Attribute VB_Name = "ExportS"
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
Public Const FILE_PREFIX As String = "S_"
' ===

' === ここから下はフィルタを使うときだけ書き換える ===
'
' フィルタを Access 側で適用するか（True=フィルタ済みエクスポート / False=生データ出力）
'   True にすると Python 側の [FILTER] セクションと同じ絞り込みを Access 側で行ってから
'   Excel へ書き出す。Python 側で [FILTER] を空にして Access に寄せたいときや、Access から
'   直接集計用ファイルを作るときに使う。False のときは従来どおり SOURCE_NAME をそのまま出す。
Public Const APPLY_FILTER As Boolean = False
'
' 種別（PLAN_COLUMN）を前方一致で絞り込む列名
Public Const PLAN_COLUMN As String = "種別"
' 種別で残す値の接頭辞群（カンマ区切り。Python の PLAN_PREFIXES に対応）
Public Const PLAN_PREFIXES As String = ""
'
' 状態（KIND_COLUMN）を完全一致で絞り込む列名
Public Const KIND_COLUMN As String = "状態"
' 状態で残す値群（カンマ区切り。Python の KINDS に対応）
Public Const KIND_VALUES As String = "完了,予定"
'
' 対象月判定に使う日付列名
Public Const DATE_COLUMN As String = "予定日"
' 対象月（'yyyy-mm' 形式）。空文字なら Format(Date, "yyyy-mm") を実行日ベースに使う
Public Const TARGET_MONTH As String = ""
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
    '    APPLY_FILTER=True のときは一時クエリ _TempExportQuery を経由して
    '    Python の [FILTER] と同じ WHERE 句で絞り込んだ結果を出す。
    DoCmd.TransferSpreadsheet _
        TransferType:=acExport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        TableName:=GetExportSource(), _
        FileName:=outputPath, _
        HasFieldNames:=True

    ' 3. 完了通知（手動実行時は MsgBox、ログ用に Debug.Print も出す）
    Debug.Print "Exported: " & outputPath
    MsgBox "エクスポート完了: " & outputPath, vbInformation, "ExportV3"

    Exit Sub

Err_ExportV3:
    ' 一時クエリが残っていると次回 CreateQueryDef がエラーになるので、
    ' エラー時も必ず削除する。
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "_TempExportQuery"
    On Error GoTo 0
    MsgBox "エクスポートに失敗しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, _
           vbCritical, "ExportV3 Error"
    Debug.Print "ExportV3 failed: " & Err.Number & " - " & Err.Description
    Resume Exit_ExportV3

Exit_ExportV3:
End Sub

' PLAN_PREFIXES / KIND_VALUES / TARGET_MONTH を SQL の WHERE 句に組み立てる。
' Python 側の [FILTER] セクションと同じ絞り込みを Access 側で再現するためのもの:
'   - PLAN_PREFIXES: PLAN_COLUMN が値のいずれかで前方一致（OR）
'   - KIND_VALUES:   KIND_COLUMN が値のいずれかと完全一致（OR）
'   - TARGET_MONTH:  DATE_COLUMN を 'yyyy-mm' に整形して一致
' 各条件は AND で結合する。列名・文字列リテラルは [...] / '...' で正しく囲む。
Private Function BuildFilterSQL() As String
    Dim parts As String
    parts = ""

    ' PLAN_PREFIXES（前方一致 OR）
    Dim prefixes() As String
    prefixes = Split(PLAN_PREFIXES, ",")
    Dim prefixClauses As String
    prefixClauses = ""
    Dim i As Long
    For i = LBound(prefixes) To UBound(prefixes)
        Dim prefix As String
        prefix = Trim$(prefixes(i))
        If Len(prefix) > 0 Then
            If Len(prefixClauses) > 0 Then prefixClauses = prefixClauses & " OR "
            prefixClauses = prefixClauses & _
                "[" & PLAN_COLUMN & "] LIKE '" & Replace(prefix, "'", "''") & "*'"
        End If
    Next i
    If Len(prefixClauses) > 0 Then
        If Len(parts) > 0 Then parts = parts & " AND "
        parts = parts & "(" & prefixClauses & ")"
    End If

    ' KIND_VALUES（完全一致 OR）
    Dim kinds() As String
    kinds = Split(KIND_VALUES, ",")
    Dim kindClauses As String
    kindClauses = ""
    For i = LBound(kinds) To UBound(kinds)
        Dim kind As String
        kind = Trim$(kinds(i))
        If Len(kind) > 0 Then
            If Len(kindClauses) > 0 Then kindClauses = kindClauses & " OR "
            kindClauses = kindClauses & _
                "[" & KIND_COLUMN & "]='" & Replace(kind, "'", "''") & "'"
        End If
    Next i
    If Len(kindClauses) > 0 Then
        If Len(parts) > 0 Then parts = parts & " AND "
        parts = parts & "(" & kindClauses & ")"
    End If

    ' TARGET_MONTH（'yyyy-mm' 一致）。空文字なら実行日ベース
    Dim targetMonth As String
    targetMonth = TARGET_MONTH
    If Len(targetMonth) = 0 Then targetMonth = Format(Date, "yyyy-mm")
    If Len(targetMonth) > 0 Then
        If Len(parts) > 0 Then parts = parts & " AND "
        parts = parts & _
            "Format([" & DATE_COLUMN & "],'yyyy-mm')='" & targetMonth & "'"
    End If

    BuildFilterSQL = "SELECT * FROM [" & SOURCE_NAME & "] WHERE " & parts
End Function

' TransferSpreadsheet に渡すテーブル/クエリ名を返す。
' APPLY_FILTER=False のときは SOURCE_NAME をそのまま使う（既存挙動と完全同一）。
' APPLY_FILTER=True のときは BuildFilterSQL() の WHERE 句を一時クエリに格納して
' その名前を返す。エクスポート後に _TempExportQuery を削除する必要はないが、
' 次回起動時に古いクエリが残っていると CreateQueryDef がエラーになるため、
' 念のため Delete を試みてから再作成する。
Private Function GetExportSource() As String
    Const TEMP_NAME As String = "_TempExportQuery"
    If APPLY_FILTER Then
        On Error Resume Next
        CurrentDb.QueryDefs.Delete TEMP_NAME
        On Error GoTo 0
        CurrentDb.CreateQueryDef TEMP_NAME, BuildFilterSQL()
        GetExportSource = TEMP_NAME
    Else
        GetExportSource = SOURCE_NAME
    End If
End Function
