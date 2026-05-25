# codex-multi-review final output

以下は、月次ブックの標準モジュールへ貼り付けて使うVBAです。

- 入口プロシージャ名: `TransferDailyClaimsToMonthly`
- 配置場所: 月次ブックの標準モジュール
- 実行方法: 月次ブックを開き、VBAエディターで標準モジュールへ貼り付けたうえで `TransferDailyClaimsToMonthly` を実行します。
- 参照設定: 追加参照は不要です。辞書は `CreateObject("Scripting.Dictionary")` のlate bindingを使います。
- 日次ファイルの場所: 月次ブックと同じフォルダ配下の `daily/2605/`
- ログ出力: `処理ログ` シートを作成または初期化し、処理結果、スキップ、エラーを一覧で記録します。
- 再実行安全性: 転記前に `月次集計` シートの `E:AJ` をクリアしてから日次ファイルを再集計するため、同じブックで再実行しても二重加算しません。

```vba
Option Explicit

Private Const TARGET_YYMM As String = "2605"
Private Const MONTHLY_SHEET_NAME As String = "月次集計"
Private Const DAILY_SHEET_NAME As String = "日次集計"
Private Const LOG_SHEET_NAME As String = "処理ログ"

Public Sub TransferDailyClaimsToMonthly()
    Dim monthlyWs As Worksheet
    Dim logWs As Worksheet
    Dim monthlyRows As Object
    Dim dailyFolder As String
    Dim fileName As String
    Dim fullPath As String
    Dim lastMonthlyRow As Long
    Dim processedFiles As Long
    Dim transferredRows As Long
    Dim skippedRows As Long
    Dim errorCount As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    On Error GoTo FatalError

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set logWs = PrepareLogSheet(ThisWorkbook)
    WriteLog logWs, "INFO", "", "", "処理開始"

    Set monthlyWs = GetWorksheet(ThisWorkbook, MONTHLY_SHEET_NAME)
    If monthlyWs Is Nothing Then
        WriteLog logWs, "ERROR", "", "", "月次シートが見つかりません: " & MONTHLY_SHEET_NAME
        GoTo CleanExit
    End If

    lastMonthlyRow = GetLastRow(monthlyWs, "A")
    If lastMonthlyRow < 2 Then
        WriteLog logWs, "ERROR", "", "", "月次シートに転記対象行がありません。"
        GoTo CleanExit
    End If

    monthlyWs.Range("E2:AJ" & lastMonthlyRow).ClearContents
    WriteLog logWs, "INFO", "", "", "月次集計の転記範囲をクリアしました。"

    Set monthlyRows = BuildMonthlyRowMap(monthlyWs, lastMonthlyRow, logWs)

    dailyFolder = ThisWorkbook.Path & Application.PathSeparator & "daily" & Application.PathSeparator & TARGET_YYMM
    If Len(Dir(dailyFolder, vbDirectory)) = 0 Then
        WriteLog logWs, "ERROR", "", "", "日次フォルダが見つかりません: daily/" & TARGET_YYMM
        GoTo CleanExit
    End If

    fileName = Dir(dailyFolder & Application.PathSeparator & "*.xls*")
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            fullPath = dailyFolder & Application.PathSeparator & fileName
            ProcessDailyWorkbook fullPath, fileName, monthlyWs, monthlyRows, logWs, transferredRows, skippedRows, errorCount
            processedFiles = processedFiles + 1
        End If
        fileName = Dir()
    Loop

    CalculateMonthlyTotals monthlyWs, lastMonthlyRow
    WriteLog logWs, "INFO", "", "", "月合計を計算しました。"
    WriteLog logWs, "SUMMARY", "", "", "処理ファイル数=" & processedFiles & _
        ", 転記件数=" & transferredRows & _
        ", スキップ件数=" & skippedRows & _
        ", エラー件数=" & errorCount

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

FatalError:
    On Error Resume Next
    If logWs Is Nothing Then Set logWs = PrepareLogSheet(ThisWorkbook)
    WriteLog logWs, "ERROR", "", "", "想定外エラー: " & Err.Number & " " & Err.Description
    Resume CleanExit
End Sub

Private Sub ProcessDailyWorkbook(ByVal fullPath As String, ByVal fileName As String, ByVal monthlyWs As Worksheet, _
                                 ByVal monthlyRows As Object, ByVal logWs As Worksheet, _
                                 ByRef transferredRows As Long, ByRef skippedRows As Long, ByRef errorCount As Long)
    Dim dailyWb As Workbook
    Dim dailyWs As Worksheet
    Dim yymmdd As String
    Dim dayNumber As Long
    Dim targetColumn As Long
    Dim dailyTotals As Object
    Dim key As Variant
    Dim targetRow As Long

    yymmdd = ExtractYymmddFromFileName(fileName, TARGET_YYMM)
    If Len(yymmdd) = 0 Then
        WriteLog logWs, "SKIP", fileName, "", "ファイル名から対象年月日を判断できません。"
        skippedRows = skippedRows + 1
        Exit Sub
    End If

    dayNumber = CLng(Right$(yymmdd, 2))
    If dayNumber < 1 Or dayNumber > 31 Then
        WriteLog logWs, "SKIP", fileName, yymmdd, "日付が1から31の範囲外です。"
        skippedRows = skippedRows + 1
        Exit Sub
    End If
    targetColumn = 4 + dayNumber

    On Error Resume Next
    Set dailyWb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=0)
    If Err.Number <> 0 Then
        WriteLog logWs, "ERROR", fileName, yymmdd, "日次ファイルを開けません: " & Err.Description
        errorCount = errorCount + 1
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo CloseWithError

    Set dailyWs = GetWorksheet(dailyWb, DAILY_SHEET_NAME)
    If dailyWs Is Nothing Then
        WriteLog logWs, "ERROR", fileName, yymmdd, "日次シートが見つかりません: " & DAILY_SHEET_NAME
        errorCount = errorCount + 1
        GoTo CloseWorkbook
    End If

    Set dailyTotals = BuildDailyTotals(dailyWs, fileName, yymmdd, logWs, skippedRows)

    For Each key In dailyTotals.Keys
        If monthlyRows.Exists(CStr(key)) Then
            targetRow = CLng(monthlyRows(CStr(key)))
            monthlyWs.Cells(targetRow, targetColumn).Value = CDbl(monthlyWs.Cells(targetRow, targetColumn).Value) + CDbl(dailyTotals(key))
            transferredRows = transferredRows + 1
        Else
            WriteLog logWs, "SKIP", fileName, yymmdd, "月次表にないキーのためスキップ: " & Replace(CStr(key), Chr$(30), " / ")
            skippedRows = skippedRows + 1
        End If
    Next key

    WriteLog logWs, "INFO", fileName, yymmdd, "日次ファイルを処理しました。"

CloseWorkbook:
    dailyWb.Close SaveChanges:=False
    Exit Sub

CloseWithError:
    WriteLog logWs, "ERROR", fileName, yymmdd, "日次ファイル処理中のエラー: " & Err.Number & " " & Err.Description
    errorCount = errorCount + 1
    On Error Resume Next
    If Not dailyWb Is Nothing Then dailyWb.Close SaveChanges:=False
End Sub

Private Function BuildDailyTotals(ByVal dailyWs As Worksheet, ByVal fileName As String, ByVal yymmdd As String, _
                                  ByVal logWs As Worksheet, ByRef skippedRows As Long) As Object
    Dim totals As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim claimValue As Variant
    Dim key As String

    Set totals = CreateObject("Scripting.Dictionary")
    totals.CompareMode = vbTextCompare

    lastRow = GetLastRow(dailyWs, "A")
    If lastRow < 2 Then
        Set BuildDailyTotals = totals
        Exit Function
    End If

    For rowIndex = 2 To lastRow
        branchCode = Trim$(CStr(dailyWs.Cells(rowIndex, "A").Value))
        businessCode = Trim$(CStr(dailyWs.Cells(rowIndex, "C").Value))
        claimValue = dailyWs.Cells(rowIndex, "E").Value

        If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
            WriteLog logWs, "SKIP", fileName, yymmdd, "支店コードまたは業務コードが空です。行=" & rowIndex
            skippedRows = skippedRows + 1
        ElseIf Not IsNumeric(claimValue) Then
            WriteLog logWs, "SKIP", fileName, yymmdd, "クレーム件数が数値ではありません。行=" & rowIndex
            skippedRows = skippedRows + 1
        Else
            key = MakeKey(branchCode, businessCode)
            If Not totals.Exists(key) Then totals.Add key, 0#
            totals(key) = CDbl(totals(key)) + CDbl(claimValue)
        End If
    Next rowIndex

    Set BuildDailyTotals = totals
End Function

Private Function BuildMonthlyRowMap(ByVal monthlyWs As Worksheet, ByVal lastMonthlyRow As Long, ByVal logWs As Worksheet) As Object
    Dim rowsByKey As Object
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim key As String

    Set rowsByKey = CreateObject("Scripting.Dictionary")
    rowsByKey.CompareMode = vbTextCompare

    For rowIndex = 2 To lastMonthlyRow
        branchCode = Trim$(CStr(monthlyWs.Cells(rowIndex, "A").Value))
        businessCode = Trim$(CStr(monthlyWs.Cells(rowIndex, "C").Value))

        If Len(branchCode) > 0 And Len(businessCode) > 0 Then
            key = MakeKey(branchCode, businessCode)
            If rowsByKey.Exists(key) Then
                WriteLog logWs, "WARN", "", "", "月次表に重複キーがあります。先の行を使用: " & branchCode & " / " & businessCode
            Else
                rowsByKey.Add key, rowIndex
            End If
        End If
    Next rowIndex

    Set BuildMonthlyRowMap = rowsByKey
End Function

Private Sub CalculateMonthlyTotals(ByVal monthlyWs As Worksheet, ByVal lastMonthlyRow As Long)
    Dim rowIndex As Long

    For rowIndex = 2 To lastMonthlyRow
        monthlyWs.Cells(rowIndex, "AJ").Value = Application.WorksheetFunction.Sum(monthlyWs.Range("E" & rowIndex & ":AI" & rowIndex))
    Next rowIndex
End Sub

Private Function PrepareLogSheet(ByVal targetWb As Workbook) As Worksheet
    Dim logWs As Worksheet

    Set logWs = GetWorksheet(targetWb, LOG_SHEET_NAME)
    If logWs Is Nothing Then
        Set logWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.Count))
        logWs.Name = LOG_SHEET_NAME
    Else
        logWs.Cells.Clear
    End If

    logWs.Range("A1:E1").Value = Array("日時", "種別", "ファイル", "対象日", "内容")
    logWs.Columns("A:E").EntireColumn.AutoFit
    Set PrepareLogSheet = logWs
End Function

Private Sub WriteLog(ByVal logWs As Worksheet, ByVal logType As String, ByVal fileName As String, ByVal targetDate As String, ByVal message As String)
    Dim nextRow As Long

    nextRow = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
    logWs.Cells(nextRow, "A").Value = Now
    logWs.Cells(nextRow, "B").Value = logType
    logWs.Cells(nextRow, "C").Value = fileName
    logWs.Cells(nextRow, "D").Value = targetDate
    logWs.Cells(nextRow, "E").Value = message
End Sub

Private Function GetWorksheet(ByVal targetWb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = targetWb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetLastRow(ByVal targetWs As Worksheet, ByVal columnLetter As String) As Long
    GetLastRow = targetWs.Cells(targetWs.Rows.Count, columnLetter).End(xlUp).Row
End Function

Private Function MakeKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeKey = Trim$(branchCode) & Chr$(30) & Trim$(businessCode)
End Function

Private Function ExtractYymmddFromFileName(ByVal fileName As String, ByVal targetYymm As String) As String
    Dim i As Long
    Dim candidate As String

    For i = 1 To Len(fileName) - 5
        candidate = Mid$(fileName, i, 6)
        If IsSixDigitText(candidate) Then
            If Left$(candidate, 4) = targetYymm Then
                ExtractYymmddFromFileName = candidate
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsSixDigitText(ByVal valueText As String) As Boolean
    Dim i As Long
    Dim codePoint As Integer

    If Len(valueText) <> 6 Then Exit Function

    For i = 1 To 6
        codePoint = Asc(Mid$(valueText, i, 1))
        If codePoint < 48 Or codePoint > 57 Then Exit Function
    Next i

    IsSixDigitText = True
End Function
```
