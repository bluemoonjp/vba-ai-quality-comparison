Option Explicit

Private Const TARGET_MONTH As String = "2605"
Private Const MONTHLY_SHEET_NAME As String = "月次集計"
Private Const DAILY_SHEET_NAME As String = "日次集計"
Private Const LOG_SHEET_NAME As String = "処理ログ"

Public Sub RunClaimsMonthlyTransfer()
    Dim wbMonthly As Workbook
    Dim wsMonthly As Worksheet
    Dim wsLog As Worksheet
    Dim monthlyMap As Object
    Dim dailyFolder As String
    Dim dailyFile As String
    Dim dayNumber As Long
    Dim processedCount As Long

    On Error GoTo FatalError

    Set wbMonthly = ThisWorkbook
    Set wsMonthly = wbMonthly.Worksheets(MONTHLY_SHEET_NAME)
    Set wsLog = PrepareLogSheet(wbMonthly)
    Set monthlyMap = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    WriteLog wsLog, "INFO", "開始", "月次クレーム集計の転記を開始しました。対象月=" & TARGET_MONTH

    ClearMonthlyData wsMonthly, wsLog
    BuildMonthlyKeyMap wsMonthly, monthlyMap, wsLog

    dailyFolder = wbMonthly.Path & Application.PathSeparator & "daily" & Application.PathSeparator & TARGET_MONTH & Application.PathSeparator
    If Len(Dir(dailyFolder, vbDirectory)) = 0 Then
        WriteLog wsLog, "ERROR", dailyFolder, "日次ファイルのフォルダが見つかりません。"
        GoTo CleanExit
    End If

    dailyFile = Dir(dailyFolder & "クレーム集計" & TARGET_MONTH & "*.xlsx")
    Do While Len(dailyFile) > 0
        dayNumber = ExtractDayNumber(dailyFile, TARGET_MONTH)
        If dayNumber >= 1 And dayNumber <= 31 Then
            ProcessDailyWorkbook dailyFolder & dailyFile, dailyFile, dayNumber, wsMonthly, monthlyMap, wsLog
            processedCount = processedCount + 1
        Else
            WriteLog wsLog, "WARN", dailyFile, "ファイル名から日付を判断できないためスキップしました。"
        End If
        dailyFile = Dir()
    Loop

    CalculateMonthlyTotals wsMonthly, wsLog
    WriteLog wsLog, "INFO", "完了", "転記が完了しました。処理ファイル数=" & CStr(processedCount)

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

FatalError:
    On Error Resume Next
    If wsLog Is Nothing Then
        Set wsLog = PrepareLogSheet(ThisWorkbook)
    End If
    WriteLog wsLog, "ERROR", "致命的エラー", Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

Private Sub ClearMonthlyData(ByVal wsMonthly As Worksheet, ByVal wsLog As Worksheet)
    Dim lastRow As Long

    lastRow = LastUsedRow(wsMonthly, "A")
    If lastRow < 2 Then
        WriteLog wsLog, "WARN", MONTHLY_SHEET_NAME, "月次集計表にデータ行がありません。"
        Exit Sub
    End If

    wsMonthly.Range("E2:AJ" & lastRow).ClearContents
    WriteLog wsLog, "INFO", MONTHLY_SHEET_NAME, "日別列と月合計をクリアしました。対象行=2:" & CStr(lastRow)
End Sub

Private Sub BuildMonthlyKeyMap(ByVal wsMonthly As Worksheet, ByVal monthlyMap As Object, ByVal wsLog As Worksheet)
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim key As String

    lastRow = LastUsedRow(wsMonthly, "A")
    For rowIndex = 2 To lastRow
        key = MakeKey(wsMonthly.Cells(rowIndex, "A").Value, wsMonthly.Cells(rowIndex, "C").Value)
        If Len(key) > 1 Then
            If monthlyMap.Exists(key) Then
                WriteLog wsLog, "WARN", "月次行 " & CStr(rowIndex), "支店コード+業務コードが月次表内で重複しています。先に見つかった行を使います。キー=" & key
            Else
                monthlyMap.Add key, rowIndex
            End If
        End If
    Next rowIndex

    WriteLog wsLog, "INFO", MONTHLY_SHEET_NAME, "月次表のキーを読み込みました。件数=" & CStr(monthlyMap.Count)
End Sub

Private Sub ProcessDailyWorkbook( _
    ByVal filePath As String, _
    ByVal fileName As String, _
    ByVal dayNumber As Long, _
    ByVal wsMonthly As Worksheet, _
    ByVal monthlyMap As Object, _
    ByVal wsLog As Worksheet)

    Dim wbDaily As Workbook
    Dim wsDaily As Worksheet
    Dim dailyTotals As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim key As String
    Dim claimCount As Double
    Dim targetRow As Long
    Dim targetColumn As Long
    Dim itemKey As Variant

    On Error GoTo DailyError

    Set dailyTotals = CreateObject("Scripting.Dictionary")
    Set wbDaily = Workbooks.Open(Filename:=filePath, UpdateLinks:=False, ReadOnly:=True)
    Set wsDaily = wbDaily.Worksheets(DAILY_SHEET_NAME)

    lastRow = LastUsedRow(wsDaily, "A")
    For rowIndex = 2 To lastRow
        key = MakeKey(wsDaily.Cells(rowIndex, "A").Value, wsDaily.Cells(rowIndex, "C").Value)
        If Len(key) > 1 Then
            If IsNumeric(wsDaily.Cells(rowIndex, "E").Value) Then
                claimCount = CDbl(wsDaily.Cells(rowIndex, "E").Value)
                If dailyTotals.Exists(key) Then
                    dailyTotals(key) = dailyTotals(key) + claimCount
                Else
                    dailyTotals.Add key, claimCount
                End If
            Else
                WriteLog wsLog, "WARN", fileName & " 行" & CStr(rowIndex), "クレーム件数が数値ではないためスキップしました。"
            End If
        Else
            WriteLog wsLog, "WARN", fileName & " 行" & CStr(rowIndex), "支店コードまたは業務コードが空のためスキップしました。"
        End If
    Next rowIndex

    targetColumn = 4 + dayNumber
    For Each itemKey In dailyTotals.Keys
        If monthlyMap.Exists(CStr(itemKey)) Then
            targetRow = CLng(monthlyMap(CStr(itemKey)))
            wsMonthly.Cells(targetRow, targetColumn).Value = dailyTotals(itemKey)
        Else
            WriteLog wsLog, "WARN", fileName, "月次表に存在しないキーのため転記しませんでした。キー=" & CStr(itemKey)
        End If
    Next itemKey

    WriteLog wsLog, "INFO", fileName, "日次ファイルを処理しました。日=" & CStr(dayNumber) & " キー数=" & CStr(dailyTotals.Count)

DailyCleanup:
    On Error Resume Next
    If Not wbDaily Is Nothing Then
        wbDaily.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Exit Sub

DailyError:
    WriteLog wsLog, "ERROR", fileName, Err.Number & ": " & Err.Description
    Resume DailyCleanup
End Sub

Private Sub CalculateMonthlyTotals(ByVal wsMonthly As Worksheet, ByVal wsLog As Worksheet)
    Dim lastRow As Long
    Dim rowIndex As Long

    lastRow = LastUsedRow(wsMonthly, "A")
    For rowIndex = 2 To lastRow
        wsMonthly.Cells(rowIndex, "AJ").Value = Application.Sum(wsMonthly.Range("E" & rowIndex & ":AI" & rowIndex))
    Next rowIndex

    WriteLog wsLog, "INFO", MONTHLY_SHEET_NAME, "月合計を再計算しました。対象行=2:" & CStr(lastRow)
End Sub

Private Function PrepareLogSheet(ByVal wb As Workbook) As Worksheet
    Dim wsLog As Worksheet

    On Error Resume Next
    Set wsLog = wb.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsLog.Name = LOG_SHEET_NAME
    Else
        wsLog.Cells.Clear
    End If

    wsLog.Range("A1:D1").Value = Array("日時", "レベル", "対象", "内容")
    wsLog.Columns("A:D").ColumnWidth = 24
    Set PrepareLogSheet = wsLog
End Function

Private Sub WriteLog(ByVal wsLog As Worksheet, ByVal levelName As String, ByVal targetName As String, ByVal messageText As String)
    Dim nextRow As Long

    If wsLog Is Nothing Then
        Exit Sub
    End If

    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Cells(nextRow, "A").Value = Now
    wsLog.Cells(nextRow, "B").Value = levelName
    wsLog.Cells(nextRow, "C").Value = targetName
    wsLog.Cells(nextRow, "D").Value = messageText
End Sub

Private Function MakeKey(ByVal branchCode As Variant, ByVal businessCode As Variant) As String
    MakeKey = Trim$(CStr(branchCode)) & "|" & Trim$(CStr(businessCode))
End Function

Private Function ExtractDayNumber(ByVal fileName As String, ByVal targetMonth As String) As Long
    Dim monthPosition As Long
    Dim dayText As String

    monthPosition = InStr(1, fileName, targetMonth, vbTextCompare)
    If monthPosition = 0 Then
        ExtractDayNumber = 0
        Exit Function
    End If

    dayText = Mid$(fileName, monthPosition + Len(targetMonth), 2)
    If IsNumeric(dayText) Then
        ExtractDayNumber = CLng(dayText)
    Else
        ExtractDayNumber = 0
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
End Function
