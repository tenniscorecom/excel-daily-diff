Attribute VB_Name = "AccessKS"
Option Explicit

' Access K/S 用集計パイプライン（CEO 直書き、2026-09-09 整理）
'
' K 用・S 用の 2 セットを集約。共通化なし（ユーザー指示）、Public Const は
' K_ / S_ プレフィックスで別名化。K 用 RunDiffK + S 用 RunDiffS。
' エクスポート機能なし（手作業で実行）。結果表示は Access のローカルテーブル
' tblResultK / tblResultS に INSERT。
'
' 起動方法: VBE で Sub RunDiffK または RunDiffS にカーソル → F5、または
'           Immediate Window で Call RunDiffK / Call RunDiffS。

' === K 用 ===
Public Const K_DB_PATH As String = "C:\作業\database.accdb"
Public Const K_SOURCE_NAME As String = "Q_K一覧"
Public Const K_OUTPUT_FOLDER As String = "C:\作業\input"
Public Const K_FILE_PREFIX As String = "K_"
Public Const K_PLAN_COLUMN As String = "種別"
Public Const K_PLAN_PREFIXES As String = ""
Public Const K_KIND_COLUMN As String = "状態"
Public Const K_KIND_VALUES As String = "完了,予定"
Public Const K_DATE_COLUMN As String = "予定日"
Public Const K_TARGET_MONTH As String = ""
Public Const K_DATE_PATTERN As String = "(20\d{6})"
Public Const K_TEMP_TABLE_PREFIX As String = "_K_"
Public Const K_KEY_COLUMN As String = "顧客番号"
Public Const K_ROLLING_MODE As Boolean = False
Public Const K_ROLLING_WINDOW_DAYS As Long = 7
Public Const K_INCREMENTAL_SAVE As Boolean = True
Public Const K_CLEANUP_TEMP_TABLES As Boolean = True
Public Const K_RESULT_TABLE As String = "tblResultK"
' ===

' === S 用 ===
Public Const S_DB_PATH As String = "C:\作業\database.accdb"
Public Const S_SOURCE_NAME As String = "Q_S一覧"
Public Const S_OUTPUT_FOLDER As String = "C:\作業\input"
Public Const S_FILE_PREFIX As String = "S_"
Public Const S_PLAN_COLUMN As String = "種別"
Public Const S_PLAN_PREFIXES As String = ""
Public Const S_KIND_COLUMN As String = "状態"
Public Const S_KIND_VALUES As String = "完了,予定"
Public Const S_DATE_COLUMN As String = "予定日"
Public Const S_TARGET_MONTH As String = ""
Public Const S_DATE_PATTERN As String = "(20\d{6})"
Public Const S_TEMP_TABLE_PREFIX As String = "_S_"
Public Const S_KEY_COLUMN As String = "顧客番号"
Public Const S_ROLLING_MODE As Boolean = False
Public Const S_ROLLING_WINDOW_DAYS As Long = 7
Public Const S_INCREMENTAL_SAVE As Boolean = True
Public Const S_CLEANUP_TEMP_TABLES As Boolean = True
Public Const S_RESULT_TABLE As String = "tblResultS"
' ===

' ====================================================
' K 用 Functions
' ====================================================

Public Sub RunDiffK()
    Dim files As Collection
    Dim entry As Variant
    Dim byRow As Object, comparedDates As Object, targetDatesByMonth As Object
    Dim dayCounts As Object
    Dim todayDate As Date, rangeStart As Date, rangeEnd As Date
    Dim fileDate As Date, businessDate As Date, lastCompared As Date
    Dim firstIndex As Long, lastIndex As Long, i As Long, comparedCount As Long
    Dim currTable As String, prevTable As String, prevName As String
    Dim errNumber As Long, errDescription As String

    On Error GoTo Err_RunDiffK
    todayDate = Date

    Set files = GetDatedFilesK()
    If files.Count = 0 Then
        MsgBox "入力フォルダに対象ファイルがありません: " & K_OUTPUT_FOLDER, _
               vbExclamation, "AccessKS (K)"
        Exit Sub
    End If

    rangeStart = RangeFloorK(todayDate)
    entry = files(files.Count)
    rangeEnd = CDate(entry(0)) - 1
    If rangeEnd > todayDate Then rangeEnd = todayDate
    If rangeEnd < rangeStart Then
        MsgBox "集計範囲に業務日がありません（K）", vbExclamation, "AccessKS (K)"
        Exit Sub
    End If

    For i = 1 To files.Count
        entry = files(i)
        fileDate = CDate(entry(0))
        If fileDate >= rangeStart + 1 And fileDate <= rangeEnd + 1 Then
            If firstIndex = 0 Then firstIndex = i
            lastIndex = i
        End If
    Next i
    If firstIndex = 0 Then
        MsgBox "読み込み対象の範囲（K）にファイルがありません", vbExclamation, "AccessKS (K)"
        Exit Sub
    End If
    If firstIndex = 1 Then
        MsgBox "K 用: 比較相手が範囲外にも存在しません", vbExclamation, "AccessKS (K)"
        Exit Sub
    End If
    firstIndex = firstIndex - 1

    Debug.Print "[K] 入力フォルダ: " & K_OUTPUT_FOLDER & "（接頭辞: " & K_FILE_PREFIX & "*.xlsx）"
    Debug.Print "[K] 読み込み範囲（業務日）: " & DateKeyK(rangeStart) & " 〜 " & DateKeyK(rangeEnd)

    Set byRow = CreateObject("Scripting.Dictionary")
    Set comparedDates = CreateObject("Scripting.Dictionary")
    Set targetDatesByMonth = CreateObject("Scripting.Dictionary")

    For i = firstIndex To lastIndex
        entry = files(i)
        fileDate = CDate(entry(0))
        currTable = K_TEMP_TABLE_PREFIX & Format$(fileDate, "yyyymmdd")
        If ImportV3FileK(CStr(entry(1)), currTable) Then
            If Len(prevTable) > 0 Then
                businessDate = fileDate - 1
                Set dayCounts = ComputeDayCountsK(prevTable, currTable, businessDate)
                MergeCountsK byRow, dayCounts, businessDate
                MarkTargetLabelsK targetDatesByMonth, businessDate
                comparedDates(DateKeyK(businessDate)) = True
                lastCompared = businessDate
                comparedCount = comparedCount + 1
                If K_INCREMENTAL_SAVE Then
                    SaveToAccessTableK byRow, DateColumnsK(rangeStart, lastCompared), _
                                          comparedDates, targetDatesByMonth
                End If
            End If
            If K_CLEANUP_TEMP_TABLES And Len(prevTable) > 0 Then DropTableIfExistsK prevTable
            prevTable = currTable
            prevName = BaseNameK(CStr(entry(1)))
        Else
            prevTable = ""
            prevName = ""
        End If
    Next i

    If comparedCount = 0 Then
        MsgBox "K 用: 比較できた業務日がありません", vbExclamation, "AccessKS (K)"
    Else
        SaveToAccessTableK byRow, DateColumnsK(rangeStart, lastCompared), comparedDates, targetDatesByMonth
        Debug.Print "[K] 結果テーブル作成: " & K_RESULT_TABLE
    End If
    If K_CLEANUP_TEMP_TABLES Then CleanUpTempTablesK
    If comparedCount > 0 Then
        MsgBox "K 用 集計完了: " & K_RESULT_TABLE, vbInformation, "AccessKS (K)"
    End If
    Exit Sub

Err_RunDiffK:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If K_CLEANUP_TEMP_TABLES Then CleanUpTempTablesK
    On Error GoTo 0
    MsgBox "K 用 集計に失敗。" & vbCrLf & "エラー: " & errNumber & vbCrLf & errDescription, vbCritical
End Sub

Private Function GetDatedFilesK() As Collection
    Dim result As Collection
    Dim fso As Object, re As Object, file As Object
    Dim dates() As Date, paths() As String
    Dim keyDate As Date, keyPath As String
    Dim fileName As String, dateText As String
    Dim fileCount As Long, i As Long, j As Long

    Set result = New Collection
    Set GetDatedFilesK = result
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(K_OUTPUT_FOLDER) Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = K_DATE_PATTERN
    re.IgnoreCase = False

    ReDim dates(0 To 0)
    ReDim paths(0 To 0)
    For Each file In fso.GetFolder(K_OUTPUT_FOLDER).Files
        fileName = file.Name
        If LCase$(Left$(fileName, Len(K_FILE_PREFIX))) = LCase$(K_FILE_PREFIX) Then
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

Private Function ImportV3FileK(ByVal filePath As String, ByVal tableName As String) As Boolean
    On Error GoTo Err_ImportK
    DropTableIfExistsK tableName
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tableName, filePath, True
    ImportV3FileK = True
    Exit Function
Err_ImportK:
    ImportV3FileK = False
End Function

Private Function ComputeDayCountsK(ByVal prevTable As String, ByVal currTable As String, _
                                    ByVal businessDate As Date) As Object
    Dim counts As Object
    Dim months As Variant
    Dim label As String
    Dim i As Long

    Set counts = CreateObject("Scripting.Dictionary")
    months = TargetMonthsK(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndexK(i)
        AddDiffCountsK counts, label, "積み上げ", currTable, prevTable, CStr(months(i))
        AddDiffCountsK counts, label, "延期", prevTable, currTable, CStr(months(i))
    Next i
    Set ComputeDayCountsK = counts
End Function

Private Sub AddDiffCountsK(counts As Object, ByVal label As String, ByVal status As String, _
                            ByVal sourceTable As String, ByVal otherTable As String, _
                            ByVal targetMonth As String)
    Dim rs As Object, db As Object
    Dim whereClause As String, sql As String, prefix As String, rowKey As String

    whereClause = BuildFilterWhereK(targetMonth)
    sql = "SELECT s.[" & K_PLAN_COLUMN & "] AS plan_value, COUNT(*) AS cnt FROM " & _
          "(SELECT [" & K_KEY_COLUMN & "], [" & K_PLAN_COLUMN & "] FROM [" & sourceTable & "]" & _
          " WHERE " & whereClause & ") AS s" & _
          " LEFT JOIN (SELECT [" & K_KEY_COLUMN & "] FROM [" & otherTable & "]" & _
          " WHERE " & whereClause & ") AS o" & _
          " ON s.[" & K_KEY_COLUMN & "] = o.[" & K_KEY_COLUMN & "]" & _
          " WHERE o.[" & K_KEY_COLUMN & "] IS NULL" & _
          " GROUP BY s.[" & K_PLAN_COLUMN & "]"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rs.EOF
        prefix = PlanPrefixOfK(NzTextK(rs.Fields("plan_value").Value))
        If Len(prefix) > 0 Then
            rowKey = label & vbTab & prefix & vbTab & status
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

Private Function BuildFilterWhereK(ByVal targetMonth As String) As String
    Dim prefixes() As String, kinds() As String
    Dim parts As String, prefixClauses As String, kindClauses As String, word As String
    Dim i As Long

    prefixes = Split(K_PLAN_PREFIXES, ",")
    For i = LBound(prefixes) To UBound(prefixes)
        word = Trim$(prefixes(i))
        If Len(word) > 0 Then
            If Len(prefixClauses) > 0 Then prefixClauses = prefixClauses & " OR "
            prefixClauses = prefixClauses & _
                "[" & K_PLAN_COLUMN & "] LIKE '" & Replace(word, "'", "''") & "*'"
        End If
    Next i
    If Len(prefixClauses) > 0 Then parts = "(" & prefixClauses & ")"

    kinds = Split(K_KIND_VALUES, ",")
    For i = LBound(kinds) To UBound(kinds)
        word = Trim$(kinds(i))
        If Len(word) > 0 Then
            If Len(kindClauses) > 0 Then kindClauses = kindClauses & " OR "
            kindClauses = kindClauses & _
                "[" & K_KIND_COLUMN & "]='" & Replace(word, "'", "''") & "'"
        End If
    Next i
    If Len(kindClauses) > 0 Then
        If Len(parts) > 0 Then parts = parts & " AND "
        parts = parts & "(" & kindClauses & ")"
    End If

    If Len(parts) > 0 Then parts = parts & " AND "
    parts = parts & "Format([" & K_DATE_COLUMN & "],'yyyy-mm')='" & targetMonth & "'"
    BuildFilterWhereK = parts & " AND [" & K_KEY_COLUMN & "] IS NOT NULL"
End Function

Private Function PlanPrefixOfK(ByVal planValue As String) As String
    Dim prefixes() As String
    Dim prefix As String
    Dim i As Long

    prefixes = Split(K_PLAN_PREFIXES, ",")
    For i = LBound(prefixes) To UBound(prefixes)
        prefix = Trim$(prefixes(i))
        If Len(prefix) > 0 Then
            If Left$(planValue, Len(prefix)) = prefix Then
                PlanPrefixOfK = prefix
                Exit Function
            End If
        End If
    Next i
    PlanPrefixOfK = ""
End Function

Private Function TargetMonthsK(ByVal businessDate As Date) As Variant
    If K_ROLLING_MODE Then
        TargetMonthsK = Array(Format$(businessDate, "yyyy-mm"), _
                               Format$(DateAdd("m", 1, businessDate), "yyyy-mm"))
    Else
        TargetMonthsK = Array(Format$(businessDate, "yyyy-mm"))
    End If
End Function

Private Function LabelForMonthIndexK(ByVal index As Long) As String
    If index = 0 Then
        LabelForMonthIndexK = "当月"
    ElseIf index = 1 Then
        LabelForMonthIndexK = "来月"
    End If
End Function

Private Function RangeFloorK(ByVal todayDate As Date) As Date
    If K_ROLLING_MODE Then
        RangeFloorK = todayDate - (K_ROLLING_WINDOW_DAYS - 1)
    Else
        RangeFloorK = DateSerial(Year(todayDate), 1, 1)
    End If
End Function

Private Sub MergeCountsK(byRow As Object, dayCounts As Object, ByVal businessDate As Date)
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
            perDate(DateKeyK(businessDate)) = dayCounts(rowKey)
        End If
    Next rowKey
End Sub

Private Sub MarkTargetLabelsK(targetDatesByMonth As Object, ByVal businessDate As Date)
    Dim months As Variant
    Dim dates As Object
    Dim label As String
    Dim i As Long

    months = TargetMonthsK(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndexK(i)
        If targetDatesByMonth.Exists(label) Then
            Set dates = targetDatesByMonth(label)
        Else
            Set dates = CreateObject("Scripting.Dictionary")
            targetDatesByMonth.Add label, dates
        End If
        dates(DateKeyK(businessDate)) = True
    Next i
End Sub

Private Function DateColumnsK(ByVal rangeStart As Date, ByVal lastCompared As Date) As Collection
    Dim result As Collection
    Dim current As Date

    Set result = New Collection
    current = rangeStart
    Do While current <= lastCompared
        result.Add current
        current = current + 1
    Loop
    Set DateColumnsK = result
End Function

Private Sub SaveToAccessTableK(byRow As Object, dates As Collection, _
                                comparedDates As Object, targetDatesByMonth As Object)
    On Error GoTo Err_SaveK
    Dim db As Object
    Set db = CurrentDb

    On Error Resume Next
    db.TableDefs.Delete K_RESULT_TABLE
    On Error GoTo 0

    Dim createSql As String
    createSql = "CREATE TABLE [" & K_RESULT_TABLE & "] ([対象月] TEXT(255), [種別] TEXT(255), [判定] TEXT(255)"
    Dim d As Variant
    For Each d In dates
        createSql = createSql & ", [" & DateKeyK(CDate(d)) & "] TEXT(255)"
    Next d
    createSql = createSql & ")"
    db.Execute createSql

    Dim prefixes() As String
    prefixes = Split(K_PLAN_PREFIXES, ",")
    Dim labels As Variant, statuses As Variant
    Dim li As Long, pi As Long, si As Long
    Dim perDate As Object, activeDates As Object
    Dim label As String, prefix As String, status As String
    Dim rowKey As String, dateText As String, cell As String

    labels = Array("当月", "来月")
    statuses = Array("積み上げ", "延期")

    For li = LBound(labels) To UBound(labels)
        label = CStr(labels(li))
        Set activeDates = Nothing
        If targetDatesByMonth.Exists(label) Then Set activeDates = targetDatesByMonth(label)
        For pi = LBound(prefixes) To UBound(prefixes)
            prefix = Trim$(prefixes(pi))
            If Len(prefix) > 0 Then
                For si = LBound(statuses) To UBound(statuses)
                    status = CStr(statuses(si))
                    rowKey = label & vbTab & prefix & vbTab & status
                    Set perDate = Nothing
                    If byRow.Exists(rowKey) Then Set perDate = byRow(rowKey)

                    Dim insertSql As String
                    insertSql = "INSERT INTO [" & K_RESULT_TABLE & "] ([対象月], [種別], [判定]"
                    Dim valuesSql As String
                    valuesSql = " VALUES ('" & Replace(label, "'", "''") & "', '" & Replace(prefix, "'", "''") & "', '" & Replace(status, "'", "''") & "'"

                    For Each d In dates
                        dateText = DateKeyK(CDate(d))
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
    Exit Sub

Err_SaveK:
    MsgBox "K 用 SaveToAccessTable エラー: " & Err.Description, vbCritical
End Sub

Private Sub DropTableIfExistsK(ByVal tableName As String)
    Dim db As Object
    Set db = CurrentDb
    On Error Resume Next
    db.TableDefs.Delete tableName
    On Error GoTo 0
    db.TableDefs.Refresh
End Sub

Private Sub CleanUpTempTablesK()
    Dim db As Object, names As Collection, tdf As Object, tableName As Variant
    Set db = CurrentDb
    Set names = New Collection
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, Len(K_TEMP_TABLE_PREFIX)) = K_TEMP_TABLE_PREFIX Then
            names.Add tdf.Name
        End If
    Next tdf
    For Each tableName In names
        DropTableIfExistsK CStr(tableName)
    Next tableName
End Sub

Private Function DateKeyK(ByVal value As Date) As String
    DateKeyK = Format$(value, "yyyy-mm-dd")
End Function

Private Function NzTextK(ByVal value As Variant) As String
    If IsNull(value) Then
        NzTextK = ""
    Else
        NzTextK = Trim$(CStr(value))
    End If
End Function

Private Function BaseNameK(ByVal path As String) As String
    Dim position As Long
    position = InStrRev(path, "\")
    If position > 0 Then
        BaseNameK = Mid$(path, position + 1)
    Else
        BaseNameK = path
    End If
End Function

' ====================================================
' S 用 Functions（K 用と 2 セット、共通化なし）
' ====================================================

Public Sub RunDiffS()
    Dim files As Collection
    Dim entry As Variant
    Dim byRow As Object, comparedDates As Object, targetDatesByMonth As Object
    Dim dayCounts As Object
    Dim todayDate As Date, rangeStart As Date, rangeEnd As Date
    Dim fileDate As Date, businessDate As Date, lastCompared As Date
    Dim firstIndex As Long, lastIndex As Long, i As Long, comparedCount As Long
    Dim currTable As String, prevTable As String, prevName As String
    Dim errNumber As Long, errDescription As String

    On Error GoTo Err_RunDiffS
    todayDate = Date

    Set files = GetDatedFilesS()
    If files.Count = 0 Then
        MsgBox "入力フォルダに対象ファイルがありません: " & S_OUTPUT_FOLDER, _
               vbExclamation, "AccessKS (S)"
        Exit Sub
    End If

    rangeStart = RangeFloorS(todayDate)
    entry = files(files.Count)
    rangeEnd = CDate(entry(0)) - 1
    If rangeEnd > todayDate Then rangeEnd = todayDate
    If rangeEnd < rangeStart Then
        MsgBox "集計範囲に業務日がありません（S）", vbExclamation, "AccessKS (S)"
        Exit Sub
    End If

    For i = 1 To files.Count
        entry = files(i)
        fileDate = CDate(entry(0))
        If fileDate >= rangeStart + 1 And fileDate <= rangeEnd + 1 Then
            If firstIndex = 0 Then firstIndex = i
            lastIndex = i
        End If
    Next i
    If firstIndex = 0 Then
        MsgBox "読み込み対象の範囲（S）にファイルがありません", vbExclamation, "AccessKS (S)"
        Exit Sub
    End If
    If firstIndex = 1 Then
        MsgBox "S 用: 比較相手が範囲外にも存在しません", vbExclamation, "AccessKS (S)"
        Exit Sub
    End If
    firstIndex = firstIndex - 1

    Debug.Print "[S] 入力フォルダ: " & S_OUTPUT_FOLDER & "（接頭辞: " & S_FILE_PREFIX & "*.xlsx）"
    Debug.Print "[S] 読み込み範囲（業務日）: " & DateKeyS(rangeStart) & " 〜 " & DateKeyS(rangeEnd)

    Set byRow = CreateObject("Scripting.Dictionary")
    Set comparedDates = CreateObject("Scripting.Dictionary")
    Set targetDatesByMonth = CreateObject("Scripting.Dictionary")

    For i = firstIndex To lastIndex
        entry = files(i)
        fileDate = CDate(entry(0))
        currTable = S_TEMP_TABLE_PREFIX & Format$(fileDate, "yyyymmdd")
        If ImportV3FileS(CStr(entry(1)), currTable) Then
            If Len(prevTable) > 0 Then
                businessDate = fileDate - 1
                Set dayCounts = ComputeDayCountsS(prevTable, currTable, businessDate)
                MergeCountsS byRow, dayCounts, businessDate
                MarkTargetLabelsS targetDatesByMonth, businessDate
                comparedDates(DateKeyS(businessDate)) = True
                lastCompared = businessDate
                comparedCount = comparedCount + 1
                If S_INCREMENTAL_SAVE Then
                    SaveToAccessTableS byRow, DateColumnsS(rangeStart, lastCompared), _
                                          comparedDates, targetDatesByMonth
                End If
            End If
            If S_CLEANUP_TEMP_TABLES And Len(prevTable) > 0 Then DropTableIfExistsS prevTable
            prevTable = currTable
            prevName = BaseNameS(CStr(entry(1)))
        Else
            prevTable = ""
            prevName = ""
        End If
    Next i

    If comparedCount = 0 Then
        MsgBox "S 用: 比較できた業務日がありません", vbExclamation, "AccessKS (S)"
    Else
        SaveToAccessTableS byRow, DateColumnsS(rangeStart, lastCompared), comparedDates, targetDatesByMonth
        Debug.Print "[S] 結果テーブル作成: " & S_RESULT_TABLE
    End If
    If S_CLEANUP_TEMP_TABLES Then CleanUpTempTablesS
    If comparedCount > 0 Then
        MsgBox "S 用 集計完了: " & S_RESULT_TABLE, vbInformation, "AccessKS (S)"
    End If
    Exit Sub

Err_RunDiffS:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If S_CLEANUP_TEMP_TABLES Then CleanUpTempTablesS
    On Error GoTo 0
    MsgBox "S 用 集計に失敗。" & vbCrLf & "エラー: " & errNumber & vbCrLf & errDescription, vbCritical
End Sub

Private Function GetDatedFilesS() As Collection
    Dim result As Collection
    Dim fso As Object, re As Object, file As Object
    Dim dates() As Date, paths() As String
    Dim keyDate As Date, keyPath As String
    Dim fileName As String, dateText As String
    Dim fileCount As Long, i As Long, j As Long

    Set result = New Collection
    Set GetDatedFilesS = result
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(S_OUTPUT_FOLDER) Then Exit Function

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = S_DATE_PATTERN
    re.IgnoreCase = False

    ReDim dates(0 To 0)
    ReDim paths(0 To 0)
    For Each file In fso.GetFolder(S_OUTPUT_FOLDER).Files
        fileName = file.Name
        If LCase$(Left$(fileName, Len(S_FILE_PREFIX))) = LCase$(S_FILE_PREFIX) Then
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

Private Function ImportV3FileS(ByVal filePath As String, ByVal tableName As String) As Boolean
    On Error GoTo Err_ImportS
    DropTableIfExistsS tableName
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tableName, filePath, True
    ImportV3FileS = True
    Exit Function
Err_ImportS:
    ImportV3FileS = False
End Function

Private Function ComputeDayCountsS(ByVal prevTable As String, ByVal currTable As String, _
                                    ByVal businessDate As Date) As Object
    Dim counts As Object
    Dim months As Variant
    Dim label As String
    Dim i As Long

    Set counts = CreateObject("Scripting.Dictionary")
    months = TargetMonthsS(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndexS(i)
        AddDiffCountsS counts, label, "積み上げ", currTable, prevTable, CStr(months(i))
        AddDiffCountsS counts, label, "延期", prevTable, currTable, CStr(months(i))
    Next i
    Set ComputeDayCountsS = counts
End Function

Private Sub AddDiffCountsS(counts As Object, ByVal label As String, ByVal status As String, _
                            ByVal sourceTable As String, ByVal otherTable As String, _
                            ByVal targetMonth As String)
    Dim rs As Object, db As Object
    Dim whereClause As String, sql As String, prefix As String, rowKey As String

    whereClause = BuildFilterWhereS(targetMonth)
    sql = "SELECT s.[" & S_PLAN_COLUMN & "] AS plan_value, COUNT(*) AS cnt FROM " & _
          "(SELECT [" & S_KEY_COLUMN & "], [" & S_PLAN_COLUMN & "] FROM [" & sourceTable & "]" & _
          " WHERE " & whereClause & ") AS s" & _
          " LEFT JOIN (SELECT [" & S_KEY_COLUMN & "] FROM [" & otherTable & "]" & _
          " WHERE " & whereClause & ") AS o" & _
          " ON s.[" & S_KEY_COLUMN & "] = o.[" & S_KEY_COLUMN & "]" & _
          " WHERE o.[" & S_KEY_COLUMN & "] IS NULL" & _
          " GROUP BY s.[" & S_PLAN_COLUMN & "]"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rs.EOF
        prefix = PlanPrefixOfS(NzTextS(rs.Fields("plan_value").Value))
        If Len(prefix) > 0 Then
            rowKey = label & vbTab & prefix & vbTab & status
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

Private Function BuildFilterWhereS(ByVal targetMonth As String) As String
    Dim prefixes() As String, kinds() As String
    Dim parts As String, prefixClauses As String, kindClauses As String, word As String
    Dim i As Long

    prefixes = Split(S_PLAN_PREFIXES, ",")
    For i = LBound(prefixes) To UBound(prefixes)
        word = Trim$(prefixes(i))
        If Len(word) > 0 Then
            If Len(prefixClauses) > 0 Then prefixClauses = prefixClauses & " OR "
            prefixClauses = prefixClauses & _
                "[" & S_PLAN_COLUMN & "] LIKE '" & Replace(word, "'", "''") & "*'"
        End If
    Next i
    If Len(prefixClauses) > 0 Then parts = "(" & prefixClauses & ")"

    kinds = Split(S_KIND_VALUES, ",")
    For i = LBound(kinds) To UBound(kinds)
        word = Trim$(kinds(i))
        If Len(word) > 0 Then
            If Len(kindClauses) > 0 Then kindClauses = kindClauses & " OR "
            kindClauses = kindClauses & _
                "[" & S_KIND_COLUMN & "]='" & Replace(word, "'", "''") & "'"
        End If
    Next i
    If Len(kindClauses) > 0 Then
        If Len(parts) > 0 Then parts = parts & " AND "
        parts = parts & "(" & kindClauses & ")"
    End If

    If Len(parts) > 0 Then parts = parts & " AND "
    parts = parts & "Format([" & S_DATE_COLUMN & "],'yyyy-mm')='" & targetMonth & "'"
    BuildFilterWhereS = parts & " AND [" & S_KEY_COLUMN & "] IS NOT NULL"
End Function

Private Function PlanPrefixOfS(ByVal planValue As String) As String
    Dim prefixes() As String
    Dim prefix As String
    Dim i As Long

    prefixes = Split(S_PLAN_PREFIXES, ",")
    For i = LBound(prefixes) To UBound(prefixes)
        prefix = Trim$(prefixes(i))
        If Len(prefix) > 0 Then
            If Left$(planValue, Len(prefix)) = prefix Then
                PlanPrefixOfS = prefix
                Exit Function
            End If
        End If
    Next i
    PlanPrefixOfS = ""
End Function

Private Function TargetMonthsS(ByVal businessDate As Date) As Variant
    If S_ROLLING_MODE Then
        TargetMonthsS = Array(Format$(businessDate, "yyyy-mm"), _
                               Format$(DateAdd("m", 1, businessDate), "yyyy-mm"))
    Else
        TargetMonthsS = Array(Format$(businessDate, "yyyy-mm"))
    End If
End Function

Private Function LabelForMonthIndexS(ByVal index As Long) As String
    If index = 0 Then
        LabelForMonthIndexS = "当月"
    ElseIf index = 1 Then
        LabelForMonthIndexS = "来月"
    End If
End Function

Private Function RangeFloorS(ByVal todayDate As Date) As Date
    If S_ROLLING_MODE Then
        RangeFloorS = todayDate - (S_ROLLING_WINDOW_DAYS - 1)
    Else
        RangeFloorS = DateSerial(Year(todayDate), 1, 1)
    End If
End Function

Private Sub MergeCountsS(byRow As Object, dayCounts As Object, ByVal businessDate As Date)
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
            perDate(DateKeyS(businessDate)) = dayCounts(rowKey)
        End If
    Next rowKey
End Sub

Private Sub MarkTargetLabelsS(targetDatesByMonth As Object, ByVal businessDate As Date)
    Dim months As Variant
    Dim dates As Object
    Dim label As String
    Dim i As Long

    months = TargetMonthsS(businessDate)
    For i = LBound(months) To UBound(months)
        label = LabelForMonthIndexS(i)
        If targetDatesByMonth.Exists(label) Then
            Set dates = targetDatesByMonth(label)
        Else
            Set dates = CreateObject("Scripting.Dictionary")
            targetDatesByMonth.Add label, dates
        End If
        dates(DateKeyS(businessDate)) = True
    Next i
End Sub

Private Function DateColumnsS(ByVal rangeStart As Date, ByVal lastCompared As Date) As Collection
    Dim result As Collection
    Dim current As Date

    Set result = New Collection
    current = rangeStart
    Do While current <= lastCompared
        result.Add current
        current = current + 1
    Loop
    Set DateColumnsS = result
End Function

Private Sub SaveToAccessTableS(byRow As Object, dates As Collection, _
                                comparedDates As Object, targetDatesByMonth As Object)
    On Error GoTo Err_SaveS
    Dim db As Object
    Set db = CurrentDb

    On Error Resume Next
    db.TableDefs.Delete S_RESULT_TABLE
    On Error GoTo 0

    Dim createSql As String
    createSql = "CREATE TABLE [" & S_RESULT_TABLE & "] ([対象月] TEXT(255), [種別] TEXT(255), [判定] TEXT(255)"
    Dim d As Variant
    For Each d In dates
        createSql = createSql & ", [" & DateKeyS(CDate(d)) & "] TEXT(255)"
    Next d
    createSql = createSql & ")"
    db.Execute createSql

    Dim prefixes() As String
    prefixes = Split(S_PLAN_PREFIXES, ",")
    Dim labels As Variant, statuses As Variant
    Dim li As Long, pi As Long, si As Long
    Dim perDate As Object, activeDates As Object
    Dim label As String, prefix As String, status As String
    Dim rowKey As String, dateText As String, cell As String

    labels = Array("当月", "来月")
    statuses = Array("積み上げ", "延期")

    For li = LBound(labels) To UBound(labels)
        label = CStr(labels(li))
        Set activeDates = Nothing
        If targetDatesByMonth.Exists(label) Then Set activeDates = targetDatesByMonth(label)
        For pi = LBound(prefixes) To UBound(prefixes)
            prefix = Trim$(prefixes(pi))
            If Len(prefix) > 0 Then
                For si = LBound(statuses) To UBound(statuses)
                    status = CStr(statuses(si))
                    rowKey = label & vbTab & prefix & vbTab & status
                    Set perDate = Nothing
                    If byRow.Exists(rowKey) Then Set perDate = byRow(rowKey)

                    Dim insertSql As String
                    insertSql = "INSERT INTO [" & S_RESULT_TABLE & "] ([対象月], [種別], [判定]"
                    Dim valuesSql As String
                    valuesSql = " VALUES ('" & Replace(label, "'", "''") & "', '" & Replace(prefix, "'", "''") & "', '" & Replace(status, "'", "''") & "'"

                    For Each d In dates
                        dateText = DateKeyS(CDate(d))
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
    Exit Sub

Err_SaveS:
    MsgBox "S 用 SaveToAccessTable エラー: " & Err.Description, vbCritical
End Sub

Private Sub DropTableIfExistsS(ByVal tableName As String)
    Dim db As Object
    Set db = CurrentDb
    On Error Resume Next
    db.TableDefs.Delete tableName
    On Error GoTo 0
    db.TableDefs.Refresh
End Sub

Private Sub CleanUpTempTablesS()
    Dim db As Object, names As Collection, tdf As Object, tableName As Variant
    Set db = CurrentDb
    Set names = New Collection
    For Each tdf In db.TableDefs
        If Left$(tdf.Name, Len(S_TEMP_TABLE_PREFIX)) = S_TEMP_TABLE_PREFIX Then
            names.Add tdf.Name
        End If
    Next tdf
    For Each tableName In names
        DropTableIfExistsS CStr(tableName)
    Next tableName
End Sub

Private Function DateKeyS(ByVal value As Date) As String
    DateKeyS = Format$(value, "yyyy-mm-dd")
End Function

Private Function NzTextS(ByVal value As Variant) As String
    If IsNull(value) Then
        NzTextS = ""
    Else
        NzTextS = Trim$(CStr(value))
    End If
End Function

Private Function BaseNameS(ByVal path As String) As String
    Dim position As Long
    position = InStrRev(path, "\")
    If position > 0 Then
        BaseNameS = Mid$(path, position + 1)
    Else
        BaseNameS = path
    End If
End Function
