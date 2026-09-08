Attribute VB_Name = "DiffKS"
Option Explicit

' 延期・積み上げ集計を Access VBA 内で完結させるパイプライン。
' Python 版 (src/run.py の run()) と同じ流れを VBA 化したもので、Python が
' 入っていない PC でも Access 単体で 集計.csv を作れるようにする。
'
' 日付が2種類出てくるので混同しないこと。
'     ファイル名の日付 … いつ時点の一覧かを表す（KS_YYYYMMDD.xlsx の YYYYMMDD）
'     案件の日付      … 絞り込みにだけ使う。対象月に入っているかを見る
' 集計表の横軸は業務日（= ファイル名の日付 - 1日）。入力ファイルは「前日終了時点」の
' データなので、KS_20260825.xlsx の中身は 8/24 終了時点であり、KS_20260824.xlsx との
' 差分は「8/24 に動いたぶん」になる。
'
' 絞り込み用の定数（PLAN_COLUMN / PLAN_PREFIXES / KIND_COLUMN / KIND_VALUES /
' DATE_COLUMN）と入力フォルダ（OUTPUT_FOLDER）・接頭辞（FILE_PREFIX）は
' ExportKS モジュールの Public Const をそのまま参照する（同じ定数を2箇所に書かない）。

' === パイプライン用（VBA 内完結） ===
'
' 入力ファイルのファイル名から日付部分を抽出する正規表現パターン
'   例: KS_20260825.xlsx → 20260825
Public Const DATE_PATTERN As String = "(20\d{6})"
'
' 一時テーブル名の接頭辞（_KS_20260825 のように日付付きで作る）
Public Const TEMP_TABLE_PREFIX As String = "_KS_"
'
' 出力 CSV のファイル名（Python の [REPORT] OUTPUT_NAME に相当）
Public Const OUTPUT_NAME As String = "K_集計.csv"
'
' 前日と当日を突き合わせるキー列（Python の [SOURCE] KEY_COLUMN）
Public Const KEY_COLUMN As String = "顧客番号"
'
' ローリングモードを使うか（False=年次累積モード / True=ローリングモード）
Public Const ROLLING_MODE As Boolean = False
'
' ローリングモードの窓幅（日数）。ROLLING_MODE=False のときは無視
Public Const ROLLING_WINDOW_DAYS As Long = 7
'
' 1日ぶん終わるたびに途中保存するか
Public Const INCREMENTAL_SAVE As Boolean = True
'
' 一時テーブルを最後に必ず削除するか
Public Const CLEANUP_TEMP_TABLES As Boolean = True
' ===

' 出力 CSV の行ラベル。対象月の実際の暦月（"2026-08" 等）ではなく、業務日の属する月を
' 「当月」、翌月を「来月」と相対ラベルで表す（Python の src/diff.py と同じ）。
Private Const LABEL_CURRENT_MONTH As String = "当月"
Private Const LABEL_NEXT_MONTH As String = "来月"

Private Const STATUS_ADDED As String = "積み上げ"
Private Const STATUS_POSTPONED As String = "延期"

' 出力 CSV のキー列見出し（PLAN_COLUMN は読み取り元の列名で、別物）
Private Const COL_TARGET_MONTH As String = "対象月"
Private Const COL_PLAN As String = "種別"
Private Const COL_STATUS As String = "判定"

' 行キーを1本の文字列にするときの区切り。VBA の Dictionary はタプルをキーに
' できないので、Python の (対象月ラベル, 種別, 判定) をタブ区切りの連結文字列にする。
' タブは対象月ラベル・種別・判定のどれにも現れない。
Private Const KEY_SEP As String = vbTab

' 集計を実行して 集計.csv を作る。VBE で本 Sub 内にカーソルを置いて F5、
' または イミディエイト ウィンドウで Call RunFullPipeline。
'
' 流れは Python の src/run.py run() と対応:
'   1. 入力フォルダの KS_*.xlsx を日付順に並べる
'   2. 集計範囲（業務日の下限・上限）を決める
'   3. 範囲内ファイル + 比較相手の直前1本を選ぶ
'   4. 1本ずつ一時テーブルへインポートする
'   5. 隣り合う2本を突き合わせて業務日ごとの件数を出す
'   6. 1日ぶん終わるたびに CSV を途中保存する
'   7. 一時テーブルを後始末する
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
               vbExclamation, "DiffPipeline"
        Exit Sub
    End If

    ' 2. 集計範囲（業務日）を決める。
    '    下限: 年次累積モードは実行年の1月1日、ローリングモードは 今日 -(N-1)日
    '    上限: 最新ファイル日付 -1日（= その日ぶんまで比較できる）を、今日で打ち切る
    rangeStart = RangeFloor(todayDate)
    entry = files(files.Count)
    rangeEnd = CDate(entry(0)) - 1
    If rangeEnd > todayDate Then rangeEnd = todayDate
    If rangeEnd < rangeStart Then
        MsgBox "集計範囲に業務日がありません（" & DateKey(rangeStart) & " 〜 " & _
               DateKey(rangeEnd) & "）", vbExclamation, "DiffPipeline"
        Exit Sub
    End If

    ' 3. 範囲内のファイルを選ぶ。業務日 → ファイル日付 は +1日。
    '    範囲内の最初のファイルには比較相手が要るので、その直前1本を範囲外から足す。
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
               "）にファイルがありません: " & OUTPUT_FOLDER, vbExclamation, "DiffPipeline"
        Exit Sub
    End If
    If firstIndex = 1 Then
        MsgBox "範囲内ファイルの比較相手が範囲外にも存在しないため、" & _
               DateKey(rangeStart) & " ぶんは集計できません", vbExclamation, "DiffPipeline"
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
        If ImportKSFile(CStr(entry(1)), currTable) Then
            If Len(prevTable) > 0 Then
                ' 業務日 = ファイル日付 -1日（入力ファイルは前日終了時点のデータ）
                businessDate = fileDate - 1
                Set dayCounts = ComputeDayCounts(prevTable, currTable, businessDate)
                MergeCounts byRow, dayCounts, businessDate
                MarkTargetLabels targetDatesByMonth, businessDate
                comparedDates(DateKey(businessDate)) = True
                lastCompared = businessDate
                comparedCount = comparedCount + 1
                LogDay businessDate, dayCounts, prevName, BaseName(CStr(entry(1)))
                If INCREMENTAL_SAVE Then
                    SaveCsv byRow, DateColumns(rangeStart, lastCompared), _
                            comparedDates, targetDatesByMonth
                End If
            End If
            ' 突き合わせが済んだ前日側の一時テーブルは、その場で落として DB を膨らませない
            ' （CLEANUP_TEMP_TABLES=False のときは中身を確認できるように残す）
            If CLEANUP_TEMP_TABLES And Len(prevTable) > 0 Then DropTableIfExists prevTable
            prevTable = currTable
            prevName = BaseName(CStr(entry(1)))
        Else
            ' 読めなかったファイルは飛ばす。次の比較相手も無くなるので前日側を空にする
            prevTable = ""
            prevName = ""
        End If
    Next i

    If comparedCount = 0 Then
        MsgBox "比較できた業務日がありません（範囲内のファイルが1本だけの可能性があります）", _
               vbExclamation, "DiffPipeline"
    Else
        ' 7. 最終保存（INCREMENTAL_SAVE=False のときはここが唯一の書き出し）
        SaveCsv byRow, DateColumns(rangeStart, lastCompared), comparedDates, targetDatesByMonth
        Debug.Print "出力しました: " & OutputPath()
    End If
    If CLEANUP_TEMP_TABLES Then CleanUpTempTables
    If comparedCount > 0 Then
        MsgBox "集計完了: " & OutputPath() & vbCrLf & _
               comparedCount & " 日ぶんを比較しました", vbInformation, "DiffPipeline"
    End If
    Exit Sub

Err_RunFullPipeline:
    ' Err の内容は後続処理で消えるので先に控える
    errNumber = Err.Number
    errDescription = Err.Description
    ' 一時テーブルが残ると次回インポートが追記になってしまうので、失敗時も必ず落とす
    On Error Resume Next
    If CLEANUP_TEMP_TABLES Then CleanUpTempTables
    On Error GoTo 0
    MsgBox "集計に失敗しました。" & vbCrLf & _
           "エラー番号: " & errNumber & vbCrLf & _
           "詳細: " & errDescription, vbCritical, "DiffPipeline Error"
    Debug.Print "RunFullPipeline failed: " & errNumber & " - " & errDescription
End Sub

' 入力フォルダ（ExportKS の OUTPUT_FOLDER。このパイプラインでは読み取り元）から
' FILE_PREFIX で始まる .xlsx を集め、ファイル名の日付が古い順に並べて返す。
' 各要素は Array(ファイル日付, フルパス)（VBA の Collection は1要素に複数値を
' 入れられないため）。
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

    ' いったん配列に貯める（Collection には並べ替えが無いので、整列してから詰める）
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

    ' 日付順（古い順）に整列。ファイル数は数百なので挿入ソートで足りる
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

' V3 ファイルを一時テーブルへインポートする。成功したら True。
' Range 引数を省略しているのでシート全体を取り込む（V3 ファイルは通常1シート）。
Private Function ImportKSFile(ByVal filePath As String, ByVal tableName As String) As Boolean
    On Error GoTo Err_Import
    ' 同名テーブルが残っていると TransferSpreadsheet が追記してしまうので先に落とす
    DropTableIfExists tableName
    DoCmd.TransferSpreadsheet _
        TransferType:=acImport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        TableName:=tableName, _
        FileName:=filePath, _
        HasFieldNames:=True
    ImportKSFile = True
    Exit Function
Err_Import:
    Debug.Print "Import failed: " & filePath & " (" & Err.Description & ")"
    ImportKSFile = False
End Function

' 前日側・当日側の2つの一時テーブルから、業務日1日ぶんの件数を数える。
' 戻り値は Scripting.Dictionary（参照設定を要らなくするため遅延バインディングで
' 作るので、宣言は Object）。キーは 対象月ラベル & KEY_SEP & 種別 & KEY_SEP & 判定、
' 値は件数。0 件の組はキーに入らない。
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
        ' 積み上げ: 当日にあって前日にない（種別は当日側の値で数える）
        AddDiffCounts counts, label, STATUS_ADDED, currTable, prevTable, CStr(months(i))
        ' 延期: 前日にあって当日にない（種別は前日側の値で数える）
        AddDiffCounts counts, label, STATUS_POSTPONED, prevTable, currTable, CStr(months(i))
    Next i
    Set ComputeDayCounts = counts
End Function

' sourceTable にあって otherTable に無いキーを種別ごとに数え、counts に足す。
' キー集合の差はフィルタ済みの全種別まとめて取り、そのあと種別で振り分ける
' （Python も current/previous の顧客番号の差を取ってから種別で分けるため。
' 種別ごとに 2 回 LEFT JOIN すると、種別が変わっただけの行を延期＋積み上げに
' 二重計上してしまう）。
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

    ' CurrentDb は呼ぶたびに新しい Database を返すので、変数に受けてから使う。
    ' rs は遅延バインディングの Object なので、! ではなく Fields(...) で明示的に読む。
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do Until rs.EOF
        ' 種別の生の値（"標準A" 等）を PLAN_PREFIXES の接頭辞に畳む
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

' Python の [FILTER] セクション（= ExportKS.BuildFilterSQL）と同じ絞り込みを、
' 対象月を引数にして WHERE 句として組み立てる。ExportKS の定数をそのまま参照するので、
' 絞り込み条件の設定はあちら 1 箇所のままにできる。
'   - PLAN_PREFIXES: PLAN_COLUMN が値のいずれかで前方一致（OR）
'   - KIND_VALUES:   KIND_COLUMN が値のいずれかと完全一致（OR）
'   - targetMonth:   DATE_COLUMN を 'yyyy-mm' に整形して一致
'   - KEY_COLUMN が空の行は突合キーにならないので落とす（Python 側と同じ）
' このパイプラインは APPLY_FILTER の値によらず常に絞り込む（集計の前提条件のため）。
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

' 種別の生の値が PLAN_PREFIXES のどれで始まるかを返す。どれにも当てはまらなければ空文字。
' Python の src/source.py _plan_prefix と同じ（PLAN_PREFIXES の並び順に前方一致）。
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

' 業務日から対象月（'yyyy-mm' の配列）を決める。
' 0 番目は必ず業務日自身の月（= 当月）。ローリングモードでは 1 番目に翌月（= 来月）を
' 無条件で加える（年次累積モードでは「来月」行は構造上できるが空のまま）。
' DateAdd が年跨ぎ（12月 → 翌年1月）も正しく扱う。
Private Function TargetMonths(ByVal businessDate As Date) As Variant
    If ROLLING_MODE Then
        TargetMonths = Array(Format$(businessDate, "yyyy-mm"), _
                             Format$(DateAdd("m", 1, businessDate), "yyyy-mm"))
    Else
        TargetMonths = Array(Format$(businessDate, "yyyy-mm"))
    End If
End Function

' TargetMonths の戻り値の位置を相対ラベルに変換する（0=当月 / 1=来月）。
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

' 読み込み範囲の下限（業務日）。年次累積モードは実行年の1月1日、ローリングモードは
' 今日を含めて N 日ぶんになるよう 今日 -(N-1)日 を返す。
Private Function RangeFloor(ByVal todayDate As Date) As Date
    If ROLLING_MODE Then
        RangeFloor = todayDate - (ROLLING_WINDOW_DAYS - 1)
    Else
        RangeFloor = DateSerial(Year(todayDate), 1, 1)
    End If
End Function

' 1日ぶんの件数を byRow（行キー → (業務日 → 件数)）に足しこむ。
' 件数 0 のセルは入れない（Python と同じ。CSV 側で「比較したが 0 件」として "0" を出す）。
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

' その業務日にどのラベル（当月/来月）が対象だったかを覚える。CSV 側で
' 「比較したがそのラベルは対象外」の列を空セルにするために使う。
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

' CSV の日付列。範囲の開始から「この時点で比較が済んだ最新の業務日」まで、1日も
' 飛ばさず並べる（ファイルが無くて比較できなかった日は空セルになる）。
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

' 集計 CSV を1本書く（途中保存・最終保存の両方から呼ばれる）。
' 行は 当月/来月 × PLAN_PREFIXES の種別 × 積み上げ/延期 の固定構成
' （PLAN_PREFIXES が既定の "標準,上位" なら 8 行）。
' セルは「件数」/「0」（比較したがそのラベルで 0 件）/「空」（比較していない、
' またはそのラベルが対象外の業務日）の3通り。
Private Sub SaveCsv(byRow As Object, dates As Collection, comparedDates As Object, _
                    targetDatesByMonth As Object)
    Dim prefixes() As String
    Dim labels As Variant, statuses As Variant, d As Variant
    Dim perDate As Object, activeDates As Object
    Dim text As String, label As String, prefix As String, status As String
    Dim rowKey As String, dateText As String, cell As String
    Dim li As Long, pi As Long, si As Long

    prefixes = Split(PLAN_PREFIXES, ",")
    labels = Array(LABEL_CURRENT_MONTH, LABEL_NEXT_MONTH)
    statuses = Array(STATUS_ADDED, STATUS_POSTPONED)

    ' ヘッダ行: 対象月, 種別, 判定, <業務日...>
    text = CsvField(COL_TARGET_MONTH) & "," & CsvField(COL_PLAN) & "," & CsvField(COL_STATUS)
    For Each d In dates
        text = text & "," & CsvField(DateKey(CDate(d)))
    Next d
    text = text & vbCrLf

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
                    text = text & CsvField(label) & "," & CsvField(prefix) & "," & _
                           CsvField(status)
                    For Each d In dates
                        dateText = DateKey(CDate(d))
                        cell = ""
                        If Not perDate Is Nothing Then
                            If perDate.Exists(dateText) Then cell = CStr(perDate(dateText))
                        End If
                        If Len(cell) = 0 Then
                            ' 比較した業務日で、そのラベルが対象だった → 0 件を "0" で表す
                            If comparedDates.Exists(dateText) Then
                                If Not activeDates Is Nothing Then
                                    If activeDates.Exists(dateText) Then cell = "0"
                                End If
                            End If
                        End If
                        text = text & "," & CsvField(cell)
                    Next d
                    text = text & vbCrLf
                Next si
            End If
        Next pi
    Next li
    WriteUtf8 OutputPath(), text
End Sub

' 出力先。V3 ファイルが並ぶ OUTPUT_FOLDER に OUTPUT_NAME で置く
' （Python 側は [REPORT] OUTPUT_FOLDER に分けているが、VBA 側は定数を増やさない）。
Private Function OutputPath() As String
    OutputPath = OUTPUT_FOLDER & "\" & OUTPUT_NAME
End Function

' カンマ・引用符・改行を含むセルだけ "..." で囲む（集計表なので通常は不要だが防御的に）。
Private Function CsvField(ByVal value As String) As String
    If InStr(value, ",") > 0 Or InStr(value, """") > 0 _
       Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
        CsvField = """" & Replace(value, """", """""") & """"
    Else
        CsvField = value
    End If
End Function

' UTF-8（BOM 付き）・改行 CRLF で書き出す。Excel でそのまま開けるようにするため。
' ADODB.Stream で UTF-8 のバイト列に変換し、BOM を捨てずにそのまま含めてから
' バイナリで書く。Open For Binary は既存ファイルを切り詰めないので、先に消して開く。
Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Dim bytes() As Byte
    Dim handle As Integer

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1                 ' adTypeBinary
    stream.Position = 0             ' UTF-8 BOM を含めるため先頭から読む
    bytes = stream.Read
    stream.Close

    If Len(Dir$(path)) > 0 Then Kill path
    handle = FreeFile
    Open path For Binary Access Write As #handle
    Put #handle, 1, bytes
    Close #handle
End Sub

' 一時テーブルを黙って落とす。DoCmd.DeleteObject は削除確認ダイアログが出る設定が
' あるため、DAO の TableDefs.Delete を使う。存在しなければ何もしない。
Private Sub DropTableIfExists(ByVal tableName As String)
    Dim db As Object

    Set db = CurrentDb
    On Error Resume Next
    db.TableDefs.Delete tableName
    On Error GoTo 0
    db.TableDefs.Refresh
End Sub

' TEMP_TABLE_PREFIX で始まるテーブルを全部削除する。
' 削除するとコレクションが変化するので、先に名前だけ集めてから消す。
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

' 業務日ごとの件数を イミディエイト ウィンドウに出す（Python 側のログと同じ粒度）。
' 対象月が2つある（ローリングモード）ときは実際の暦月も添える。
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

' ログ表示用に、種別を区別せず (対象月ラベル, 判定) で合計する。
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

' 業務日を CSV の列見出し・Dictionary のキーに使う 'yyyy-mm-dd' 文字列にする。
Private Function DateKey(ByVal value As Date) As String
    DateKey = Format$(value, "yyyy-mm-dd")
End Function

' Null を空文字にして前後の空白を落とす（Python の src/source.py _text と同じ）。
Private Function NzText(ByVal value As Variant) As String
    If IsNull(value) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(value))
    End If
End Function

' フルパスからファイル名だけを取り出す（ログ表示用）。
Private Function BaseName(ByVal path As String) As String
    Dim position As Long

    position = InStrRev(path, "\")
    If position > 0 Then
        BaseName = Mid$(path, position + 1)
    Else
        BaseName = path
    End If
End Function
