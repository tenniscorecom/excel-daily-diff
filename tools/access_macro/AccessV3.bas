Attribute VB_Name = "AccessV3"
Option Explicit

' Access V3 用集計パイプライン（CEO 直書き、2026-09-09 整理）
'
' 既存の ExportV3.bas（エクスポート）と DiffPipeline.bas（集計）を 1 ファイルに統合。
' エクスポート機能なし（手作業で実行）。結果表示は Access のローカルテーブル
' tblResult に INSERT（CSV 取り込みなし、CSV と同じ見え方）。
'
' 起動方法: VBE で Sub RunFullPipeline にカーソル → F5、または
'           Immediate Window で Call RunFullPipeline。
'
' 日付が2種類出てくるので混同しないこと。
'     ファイル名の日付 … いつ時点の一覧かを表す（V3_YYYYMMDD.xlsx の YYYYMMDD）
'     案件の日付      … 絞り込みにだけ使う。対象月に入っているかを見る
' 集計表の横軸は業務日（= ファイル名の日付 - 1日）。入力ファイルは「前日終了時点」の
' データなので、V3_20260825.xlsx の中身は 8/24 終了時点であり、V3_20260824.xlsx との
' 差分は「8/24 に動いたぶん」になる。

' === ここを実行環境に合わせて書き換える ===
' Access データベースのフルパス
Public Const DB_PATH As String = "C:\作業\database.accdb"
' エクスポート対象のテーブルまたはクエリ名
Public Const SOURCE_NAME As String = "Q_一覧"
' 出力先フォルダ（既存の INPUT_FOLDER と一致させる）
Public Const OUTPUT_FOLDER As String = "C:\作業\input"
' 出力ファイル名の接頭辞
Public Const FILE_PREFIX As String = "V3_"
' ===

' === ここから下はフィルタを使うときだけ書き換える ===
Public Const APPLY_FILTER As Boolean = False
Public Const PLAN_COLUMN As String = "種別"
Public Const PLAN_PREFIXES As String = "標準,上位"
Public Const KIND_COLUMN As String = "状態"
Public Const KIND_VALUES As String = "完了,予定"
Public Const DATE_COLUMN As String = "予定日"
Public Const TARGET_MONTH As String = ""
' ===

' === パイプライン用（VBA 内完結） ===
Public Const DATE_PATTERN As String = "(20\d{6})"
Public Const TEMP_TABLE_PREFIX As String = "_V3_"
Public Const KEY_COLUMN As String = "顧客番号"
Public Const ROLLING_MODE As Boolean = False
Public Const ROLLING_WINDOW_DAYS As Long = 7
Public Const INCREMENTAL_SAVE As Boolean = True
Public Const CLEANUP_TEMP_TABLES As Boolean = True
' ===

' 結果テーブル（Access のローカルテーブル。CSV と同じ列構成）
Public Const RESULT_TABLE As String = "tblResult"

' 出力 CSV の行ラベル
Private Const LABEL_CURRENT_MONTH As String = "当月"
Private Const LABEL_NEXT_MONTH As String = "来月"
Private Const STATUS_ADDED As String = "積み上げ"
Private Const STATUS_POSTPONED As String = "延期"
Private Const COL_TARGET_MONTH As String = "対象月"
Private Const COL_PLAN As String = "種別"
Private Const COL_STATUS As String = "判定"
Private Const KEY_SEP As String = vbTab

' 集計を実行して Access のローカルテーブル tblResult に書き出す。VBE で本 Sub 内に
' カーソルを置いて F5、または イミディエイト ウィンドウで Call RunFullPipeline。
Public Sub RunFullPipeline()
    Dim files As Collection
    Dim entry As Variant
    Dim byRow As Object, comparedDates As Object, targetDatesByMonth As Object
    Dim dayCounts As Object
    Dim todayDate As Date, rangeStart As Date, rangeEnd As Date
    Dim fileDate As Date, businessDate As Date, lastCompared As Date
    Dim firstIndex As Long, lastIndex As Long, i As Long, comparedCount As Long
    Dim currTable As String, prevTable As String, prevName As String
    Dim errNumber As Long, errDescription As String

    On Error GoTo Err_RunFullPipeline
    todayDate = Date

    ' 1. 入力フォルダのファイルを日付順に並べる
    Set files = GetDatedFiles()
    If files.Count = 0 Then
        MsgBox "入力フォルダに対象ファイルがありません: " & OUTPUT_FOLDER, _
               vbExclamation, "AccessV3"
        Exit Sub
    End If

    ' 2. 集計範囲（業務日）を決める
    rangeStart = RangeFloor(todayDate)
    entry = files(files.Count)
    rangeEnd = CDate(entry(0)) - 1
    If rangeEnd > todayDate Then rangeEnd = todayDate
    If rangeEnd < rangeStart Then
        MsgBox "集計範囲に業務日がありません（" & DateKey(rangeStart) & " 〜 " & _
               DateKey(rangeEnd) & "）", vbExclamation, "AccessV3"
        Exit Sub
    End If

    ' 3. 範囲内のファイルを選ぶ。業務日 → ファイル日付 は +1日
    For i = 1 To files.Count
        entry = files(i)
        fileDate = CDate(entry(0))
        If fileDate >= rangeStart + 1 And fileDate <= rangeEnd + 1 Then
            If firstIndex = 0 Then firstIndex = i
            lastIndex = i
        End If
    Next i
    If firstIndex = 0 Then
        MsgBox "読み込み対象の範囲（" & DateKey(rangeStart) & " 〜 " & DateKey(rangeEnd) & _
               "）にファイルがありません: " & OUTPUT_FOLDER, vbExclamation, "AccessV3"
        Exit Sub
    End If
    If firstIndex = 1 Then
        MsgBox "範囲内ファイルの比較相手が範囲外にも存在しないため、" & _
               DateKey(rangeStart) & " ぶんは集計できません", vbExclamation, "AccessV3"
        Exit Sub
    End If
    firstIndex = firstIndex - 1

    Debug.Print "入力フォルダ: " & OUTPUT_FOLDER & "（接頭辞: " & FILE_PREFIX & "*.xlsx）"
    Debug.Print "実行モード: " & IIf(ROLLING_MODE, _
        "ローリング（直近" & ROLLING_WINDOW_DAYS & "日）", "年次累積")
    Debug.Print "読み込み範囲（業務日）: " & DateKey(rangeStart) & " 〜 " & DateKey(rangeEnd)
    Debug.Print "入力フォルダのファイル " & files.Count & " 件のうち、" & _
        (lastIndex - firstIndex + 1) & " 件を読み込みます（先頭1件は範囲外の比較相手）"
    If INCREMENTAL_SAVE Then Debug.Print "1日ぶんごとに保存しながら進めます"

    Set byRow = CreateObject("Scripting.Dictionary")
    Set comparedDates = CreateObject("Scripting.Dictionary")
    Set targetDatesByMonth = CreateObject("Scripting.Dictionary")

    ' 4〜6. 1本ずつインポートして、隣り合う2本を突き合わせる
    For i = firstIndex To lastIndex
        entry = files(i)
        fileDate = CDate(entry(0))
        currTable = TEMP_TABLE_PREFIX & Format$(fileDate, "yyyymmdd")
        Debug.Print "(" & (i - firstIndex + 1) & "/" & (lastIndex - firstIndex + 1) & ") " & _
            BaseName(CStr(entry(1)))
        If ImportV3File(CStr(entry(1)), currTable) Then
            If Len(prevTable) > 0 Then
                businessDate = fileDate - 1
                Set dayCounts = ComputeDayCounts(prevTable, currTable, businessDate)
                MergeCounts byRow, dayCounts, businessDate
                MarkTargetLabels targetDatesByMonth, businessDate
                comparedDates(DateKey(businessDate)) = True
                lastCompared = businessDate
                comparedCount = comparedCount + 1
                LogDay businessDate, dayCounts, prevName, BaseName(CStr(entry(1)))
                If INCREMENTAL_SAVE Then
                    SaveToAccessTable byRow, DateColumns(rangeStart, lastCompared), _
                                       comparedDates, targetDatesByMonth
                End If
            End If
            If CLEANUP_TEMP_TABLES And Len(prevTable) > 0 Then DropTableIfExists prevTable
            prevTable = currTable
            prevName = BaseName(CStr(entry(1)))
        Else
            prevTable = ""
            prevName = ""
        End If
    Next i

    If comparedCount = 0 Then
        MsgBox "比較できた業務日がありません（範囲内のファイルが1本だけの可能性があります）", _
               vbExclamation, "AccessV3"
    Else
        SaveToAccessTable byRow, DateColumns(rangeStart, lastCompared), comparedDates, targetDatesByMonth
        Debug.Print "結果テーブル作成: " & RESULT_TABLE
    End If
    If CLEANUP_TEMP_TABLES Then CleanUpTempTables
    If comparedCount > 0 Then
        MsgBox "集計完了: " & RESULT_TABLE & vbCrLf & _
               comparedCount & " 日ぶんを比較しました", vbInformation, "AccessV3"
    End If
    Exit Sub

Err_RunFullPipeline:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If CLEANUP_TEMP_TABLES Then CleanUpTempTables
    On Error GoTo 0
    MsgBox "集計に失敗しました。" & vbCrLf & _
           "エラー番号: " & errNumber & vbCrLf & _
           "詳細: " & errDescription, vbCritical, "AccessV3 Error"
    Debug.Print "RunFullPipeline failed: " & errNumber & " - " & errDescription
End Sub

Private Function GetDatedFiles() As Collection
    Dim result As Collection
    Dim fso As Object, re As Object, file As Object
    Dim dates() As Date, paths() As String
    Dim keyDate As Date, keyPath As String
    Dim fileName As String, dateText As String
    Dim fileCount As Long, i As Long, j As Long

    Set result = New Collection
    Set GetDatedFiles = result
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(OUTPUT_FOLDER) Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = DATE_PATTERN
    re.IgnoreCase = False

    ReDim dates(0 To 0)
    ReDim paths(0 To 0)
    For Each file In fso.GetFolder(OUTPUT_FOLDER).Files
        fileName = file.Name
        If LCase$(Left$(fileName, Len(FILE_PREFIX))) = LCase$(FILE_PREFIX) Then
            If LCase$(fso.GetExtensionName(fileName)) = "xlsx" Then
                If re.Test(fileName) Then
                    dateText = re.Execute(fileName)(0).SubMatches(0)
                    ReDim Preserve dates(0 To fileCount)
                    ReDim Preserve paths(0 To fileCount)
                    dates(fileCount) = DateSerial(CInt(Left$(dateText, 4)), _
                                                  CInt(Mid$(dateText, 5, 2)), _
                                                  CInt(Right$(dateText, 2)))
                    paths(fileCount) = file.Path
                    fileCount = fileCount + 1
                End If
            End If
        End If
    Next file
    If fileCount = 0 Then Exit Function

    For i = 1 To fileCount - 1
        keyDate = dates(i)
        keyPath = paths(i)
        j = i - 1
        Do While j >= 0
            If dates(j) <= keyDate Then Exit Do
            dates(j + 1) = dates(j)
            paths(j + 1) = paths(j)
            j = j - 1
        Loop
        dates(j + 1) = keyDate
        paths(j + 1) = keyPath
    Next i
    For i = 0 To fileCount - 1
        result.Add Array(dates(i), paths(i))
    Next i
End Function

Private Function ImportV3File(ByVal filePath As String, ByVal tableName As String) As Boolean
    On Error GoTo Err_Import
    DropTableIfExists tableName
    DoCmd.TransferSpreadsheet _
        TransferType:=acImport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        TableName:=tableName, _
        FileName:=filePath, _
        HasFieldNames:=True
    ImportV3File = True
    Exit Function
Err_Import:
    Debug.Print "Import failed: " & filePath & " (" & Err.Description & ")"
    ImportV3File = False
End Function

Private Function ComputeDayCounts(ByVal prevTable As String, ByVal currTable As String, _
                                  ByVal businessDate As Date) As Object
    Dim counts As Object
    Dim months As Variant
    Dim label As String
    Dim i As Long

    Set counts = CreateObject("Scripting.Dictionary")
    months = TargetMonths(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndex(i)
        AddDiffCounts counts, label, STATUS_ADDED, currTable, prevTable, CStr(months(i))
        AddDiffCounts counts, label, STATUS_POSTPONED, prevTable, currTable, CStr(months(i))
    Next i
    Set ComputeDayCounts = counts
End Function

Private Sub AddDiffCounts(counts As Object, ByVal label As String, ByVal status As String, _
                          ByVal sourceTable As String, ByVal otherTable As String, _
                          ByVal targetMonth As String)
    Dim rs As Object, db As Object
    Dim whereClause As String, sql As String, prefix As String, rowKey As String

    whereClause = BuildFilterWhere(targetMonth)
    sql = "SELECT s.[" & PLAN_COLUMN & "] AS plan_value, COUNT(*) AS cnt FROM " & _
          "(SELECT [" & KEY_COLUMN & "], [" & PLAN_COLUMN & "] FROM [" & sourceTable & "]" & _
          " WHERE " & whereClause & ") AS s" & _
          " LEFT JOIN (SELECT [" & KEY_COLUMN & "] FROM [" & otherTable & "]" & _
          " WHERE " & whereClause & ") AS o" & _
          " ON s.[" & KEY_COLUMN & "] = o.[" & KEY_COLUMN & "]" & _
          " WHERE o.[" & KEY_COLUMN & "] IS NULL" & _
          " GROUP BY s.[" & PLAN_COLUMN & "]"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rs.EOF
        prefix = PlanPrefixOf(NzText(rs.Fields("plan_value").Value))
        If Len(prefix) > 0 Then
            rowKey = label & KEY_SEP & prefix & KEY_SEP & status
            If counts.Exists(rowKey) Then
                counts(rowKey) = counts(rowKey) + CLng(rs.Fields("cnt").Value)
            Else
                counts.Add rowKey, CLng(rs.Fields("cnt").Value)
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function BuildFilterWhere(ByVal targetMonth As String) As String
    Dim prefixes() As String, kinds() As String
    Dim parts As String, prefixClauses As String, kindClauses As String, word As String
    Dim i As Long

    prefixes = Split(PLAN_PREFIXES, ",")
    For i = LBound(prefixes) To UBound(prefixes)
        word = Trim$(prefixes(i))
        If Len(word) > 0 Then
            If Len(prefixClauses) > 0 Then prefixClauses = prefixClauses & " OR "
            prefixClauses = prefixClauses & _
                "[" & PLAN_COLUMN & "] LIKE '" & Replace(word, "'", "''") & "*'"
        End If
    Next i
    If Len(prefixClauses) > 0 Then parts = "(" & prefixClauses & ")"

    kinds = Split(KIND_VALUES, ",")
    For i = LBound(kinds) To UBound(kinds)
        word = Trim$(kinds(i))
        If Len(word) > 0 Then
            If Len(kindClauses) > 0 Then kindClauses = kindClauses & " OR "
            kindClauses = kindClauses & _
                "[" & KIND_COLUMN & "]='" & Replace(word, "'", "''") & "'"
        End If
    Next i
    If Len(kindClauses) > 0 Then
        If Len(parts) > 0 Then parts = parts & " AND "
        parts = parts & "(" & kindClauses & ")"
    End If

    If Len(parts) > 0 Then parts = parts & " AND "
    parts = parts & "Format([" & DATE_COLUMN & "],'yyyy-mm')='" & targetMonth & "'"
    BuildFilterWhere = parts & " AND [" & KEY_COLUMN & "] IS NOT NULL"
End Function

Private Function PlanPrefixOf(ByVal planValue As String) As String
    Dim prefixes() As String
    Dim prefix As String
    Dim i As Long

    prefixes = Split(PLAN_PREFIXES, ",")
    For i = LBound(prefixes) To UBound(prefixes)
        prefix = Trim$(prefixes(i))
        If Len(prefix) > 0 Then
            If Left$(planValue, Len(prefix)) = prefix Then
                PlanPrefixOf = prefix
                Exit Function
            End If
        End If
    Next i
    PlanPrefixOf = ""
End Function

Private Function TargetMonths(ByVal businessDate As Date) As Variant
    If ROLLING_MODE Then
        TargetMonths = Array(Format$(businessDate, "yyyy-mm"), _
                             Format$(DateAdd("m", 1, businessDate), "yyyy-mm"))
    Else
        TargetMonths = Array(Format$(businessDate, "yyyy-mm"))
    End If
End Function

Private Function LabelForMonthIndex(ByVal index As Long) As String
    If index = 0 Then
        LabelForMonthIndex = LABEL_CURRENT_MONTH
    ElseIf index = 1 Then
        LabelForMonthIndex = LABEL_NEXT_MONTH
    Else
        Err.Raise 5, "LabelForMonthIndex", _
            "対象月は当月・来月の2つまでしか扱えません（index=" & index & "）"
    End If
End Function

Private Function RangeFloor(ByVal todayDate As Date) As Date
    If ROLLING_MODE Then
        RangeFloor = todayDate - (ROLLING_WINDOW_DAYS - 1)
    Else
        RangeFloor = DateSerial(Year(todayDate), 1, 1)
    End If
End Function

Private Sub MergeCounts(byRow As Object, dayCounts As Object, ByVal businessDate As Date)
    Dim rowKey As Variant
    Dim perDate As Object

    For Each rowKey In dayCounts.Keys
        If dayCounts(rowKey) > 0 Then
            If byRow.Exists(rowKey) Then
                Set perDate = byRow(rowKey)
            Else
                Set perDate = CreateObject("Scripting.Dictionary")
                byRow.Add rowKey, perDate
            End If
            perDate(DateKey(businessDate)) = dayCounts(rowKey)
        End If
    Next rowKey
End Sub

Private Sub MarkTargetLabels(targetDatesByMonth As Object, ByVal businessDate As Date)
    Dim months As Variant
    Dim dates As Object
    Dim label As String
    Dim i As Long

    months = TargetMonths(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndex(i)
        If targetDatesByMonth.Exists(label) Then
            Set dates = targetDatesByMonth(label)
        Else
            Set dates = CreateObject("Scripting.Dictionary")
            targetDatesByMonth.Add label, dates
        End If
        dates(DateKey(businessDate)) = True
    Next i
End Sub

Private Function DateColumns(ByVal rangeStart As Date, ByVal lastCompared As Date) As Collection
    Dim result As Collection
    Dim current As Date

    Set result = New Collection
    current = rangeStart
    Do While current <= lastCompared
        result.Add current
        current = current + 1
    Loop
    Set DateColumns = result
End Function

' 集計結果を Access のローカルテーブル RESULT_TABLE に書き出す（CSV の代わりに）。
' 既存の SaveCsv を SaveToAccessTable に変更。SQL の INSERT で 1 行ずつ書き込む。
Private Sub SaveToAccessTable(byRow As Object, dates As Collection, _
                              comparedDates As Object, targetDatesByMonth As Object)
    On Error GoTo Err_SaveToAccessTable
    Dim db As Object
    Set db = CurrentDb

    ' 既存テーブルを DROP
    On Error Resume Next
    db.TableDefs.Delete RESULT_TABLE
    On Error GoTo 0

    ' CREATE TABLE 文を構築（固定列 + 業務日列）
    Dim createSql As String
    createSql = "CREATE TABLE [" & RESULT_TABLE & "] ([" & COL_TARGET_MONTH & "] TEXT(255), [" & COL_PLAN & "] TEXT(255), [" & COL_STATUS & "] TEXT(255)"
    Dim d As Variant
    For Each d In dates
        createSql = createSql & ", [" & DateKey(CDate(d)) & "] TEXT(255)"
    Next d
    createSql = createSql & ")"
    db.Execute createSql

    ' INSERT 行
    Dim prefixes() As String
    prefixes = Split(PLAN_PREFIXES, ",")
    Dim labels As Variant, statuses As Variant
    Dim li As Long, pi As Long, si As Long
    Dim perDate As Object, activeDates As Object
    Dim label As String, prefix As String, status As String
    Dim rowKey As String, dateText As String, cell As String

    labels = Array(LABEL_CURRENT_MONTH, LABEL_NEXT_MONTH)
    statuses = Array(STATUS_ADDED, STATUS_POSTPONED)

    For li = LBound(labels) To UBound(labels)
        label = CStr(labels(li))
        Set activeDates = Nothing
        If targetDatesByMonth.Exists(label) Then Set activeDates = targetDatesByMonth(label)
        For pi = LBound(prefixes) To UBound(prefixes)
            prefix = Trim$(prefixes(pi))
            If Len(prefix) > 0 Then
                For si = LBound(statuses) To UBound(statuses)
                    status = CStr(statuses(si))
                    rowKey = label & KEY_SEP & prefix & KEY_SEP & status
                    Set perDate = Nothing
                    If byRow.Exists(rowKey) Then Set perDate = byRow(rowKey)

                    Dim insertSql As String
                    insertSql = "INSERT INTO [" & RESULT_TABLE & "] ([" & COL_TARGET_MONTH & "], [" & COL_PLAN & "], [" & COL_STATUS & "]"
                    Dim valuesSql As String
                    valuesSql = " VALUES ('" & Replace(label, "'", "''") & "', '" & Replace(prefix, "'", "''") & "', '" & Replace(status, "'", "''") & "'"

                    For Each d In dates
                        dateText = DateKey(CDate(d))
                        insertSql = insertSql & ", [" & dateText & "]"
                        cell = ""
                        If Not perDate Is Nothing Then
                            If perDate.Exists(dateText) Then cell = CStr(perDate(dateText))
                        End If
                        If Len(cell) = 0 Then
                            If comparedDates.Exists(dateText) Then
                                If Not activeDates Is Nothing Then
                                    If activeDates.Exists(dateText) Then cell = "0"
                                End If
                            End If
                        End If
                        valuesSql = valuesSql & ", '" & Replace(cell, "'", "''") & "'"
                    Next d
                    insertSql = insertSql & ")" & valuesSql & ")"

                    db.Execute insertSql
                Next si
            End If
        Next pi
    Next li

    Debug.Print "結果テーブル作成: " & RESULT_TABLE
    Exit Sub

Err_SaveToAccessTable:
    MsgBox "結果テーブルの作成に失敗しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical, "SaveToAccessTable Error"
End Sub

Private Sub DropTableIfExists(ByVal tableName As String)
    Dim db As Object

    Set db = CurrentDb
    On Error Resume Next
    db.TableDefs.Delete tableName
    On Error GoTo 0
    db.TableDefs.Refresh
End Sub

Private Sub CleanUpTempTables()
    Dim db As Object
    Dim names As Collection
    Dim tdf As Object
    Dim tableName As Variant

    Set db = CurrentDb
    Set names = New Collection
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, Len(TEMP_TABLE_PREFIX)) = TEMP_TABLE_PREFIX Then
            names.Add tdf.Name
        End If
    Next tdf
    For Each tableName In names
        DropTableIfExists CStr(tableName)
    Next tableName
End Sub

Private Sub LogDay(ByVal businessDate As Date, dayCounts As Object, _
                   ByVal prevName As String, ByVal currName As String)
    Dim months As Variant
    Dim label As String, suffix As String
    Dim i As Long

    months = TargetMonths(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndex(i)
        suffix = ""
        If UBound(months) > LBound(months) Then suffix = " [" & months(i) & "]"
        Debug.Print "業務日 " & DateKey(businessDate) & "（" & label & "）" & suffix & _
            ": 積み上げ " & SumForLabel(dayCounts, label, STATUS_ADDED) & " 件 / 延期 " & _
            SumForLabel(dayCounts, label, STATUS_POSTPONED) & " 件" & _
            "（" & prevName & " → " & currName & "）"
    Next i
End Sub

Private Function SumForLabel(dayCounts As Object, ByVal label As String, _
                             ByVal status As String) As Long
    Dim total As Long
    Dim rowKey As Variant
    Dim parts() As String

    For Each rowKey In dayCounts.Keys
        parts = Split(CStr(rowKey), KEY_SEP)
        If parts(0) = label And parts(2) = status Then total = total + dayCounts(rowKey)
    Next rowKey
    SumForLabel = total
End Function

Private Function DateKey(ByVal value As Date) As String
    DateKey = Format$(value, "yyyy-mm-dd")
End Function

Private Function NzText(ByVal value As Variant) As String
    If IsNull(value) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(value))
    End If
End Function

Private Function BaseName(ByVal path As String) As String
    Dim position As Long

    position = InStrRev(path, "\")
    If position > 0 Then
        BaseName = Mid$(path, position + 1)
    Else
        BaseName = path
    End If
End Function
