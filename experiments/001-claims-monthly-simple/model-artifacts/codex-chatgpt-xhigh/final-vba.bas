Attribute VB_Name = "modClaimsMonthlyImport"
Option Explicit

Private Const MONTHLY_SHEET_NAME As String = "月次集計"
Private Const DAILY_SHEET_NAME As String = "日次集計"
Private Const LOG_SHEET_NAME As String = "処理ログ"

Private Const FIRST_DATA_ROW As Long = 2
Private Const COL_BRANCH_CODE As Long = 1
Private Const COL_BUSINESS_CODE As Long = 3
Private Const COL_CLAIM_COUNT As Long = 5
Private Const FIRST_DAY_COL As Long = 5
Private Const LAST_DAY_COL As Long = 35
Private Const COL_MONTH_TOTAL As Long = 36

Public Sub ImportDailyClaimsToMonthly()
    Dim wbMonthly As Workbook
    Dim wsMonthly As Worksheet
    Dim wsLog As Worksheet
    Dim targetRows As Object
    Dim monthCode As String
    Dim dailyFolder As String
    Dim fileName As String
    Dim dailyDate As String
    Dim dayNumber As Long
    Dim lastMonthlyRow As Long
    Dim processedFiles As Long
    Dim importedRows As Long
    Dim skippedRows As Long
    Dim warnings As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    On Error GoTo FatalError

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wbMonthly = ThisWorkbook
    Set wsLog = PrepareLogSheet(wbMonthly)
    LogMessage wsLog, "INFO", "", "", "処理を開始しました。"

    If Len(wbMonthly.Path) = 0 Then
        Err.Raise vbObjectError + 1001, , "月次ブックを保存してから実行してください。"
    End If

    Set wsMonthly = GetWorksheet(wbMonthly, MONTHLY_SHEET_NAME)
    If wsMonthly Is Nothing Then
        Err.Raise vbObjectError + 1002, , "月次シート '" & MONTHLY_SHEET_NAME & "' が見つかりません。"
    End If

    monthCode = ExtractDigitRun(wbMonthly.Name, 4)
    If Len(monthCode) <> 4 Then
        Err.Raise vbObjectError + 1003, , "月次ブック名から対象月 YYMM を判断できません。"
    End If

    dailyFolder = JoinPath(JoinPath(wbMonthly.Path, "daily"), monthCode)
    If Not FolderExists(dailyFolder) Then
        Err.Raise vbObjectError + 1004, , "日次フォルダが見つかりません: " & dailyFolder
    End If

    lastMonthlyRow = LastUsedRow(wsMonthly, COL_BRANCH_CODE)
    If lastMonthlyRow < FIRST_DATA_ROW Then
        Err.Raise vbObjectError + 1005, , "月次集計シートにデータ行がありません。"
    End If

    Set targetRows = BuildTargetRowMap(wsMonthly, lastMonthlyRow, wsLog, warnings)
    If targetRows.Count = 0 Then
        Err.Raise vbObjectError + 1006, , "月次集計シートに有効な支店コード + 業務コードがありません。"
    End If

    ResetMonthlyValues wsMonthly, lastMonthlyRow

    fileName = Dir(JoinPath(dailyFolder, "*.xls*"))
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            dailyDate = ExtractDigitRun(fileName, 6)
            If Len(dailyDate) = 6 And Left$(dailyDate, 4) = monthCode Then
                dayNumber = CLng(Right$(dailyDate, 2))
                If dayNumber >= 1 And dayNumber <= 31 Then
                    processedFiles = processedFiles + 1
                    ProcessDailyWorkbook JoinPath(dailyFolder, fileName), fileName, dayNumber, wsMonthly, targetRows, wsLog, importedRows, skippedRows, warnings
                Else
                    warnings = warnings + 1
                    LogMessage wsLog, "WARN", fileName, "", "ファイル名の日付が1日から31日の範囲外です。"
                End If
            End If
        End If
        fileName = Dir()
    Loop

    If processedFiles = 0 Then
        warnings = warnings + 1
        LogMessage wsLog, "WARN", "", "", "対象月の日次ファイルが見つかりませんでした。"
    End If

    CalculateMonthlyTotals wsMonthly, lastMonthlyRow

    LogMessage wsLog, "INFO", "", "", "処理が完了しました。対象ファイル数=" & processedFiles & _
        ", 転記行数=" & importedRows & ", スキップ行数=" & skippedRows & ", 警告数=" & warnings
    wsLog.Columns("A:E").AutoFit

CleanExit:
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

FatalError:
    LogMessage wsLog, "ERROR", "", "", "処理を中断しました: " & Err.Description
    Resume CleanExit
End Sub

Private Function PrepareLogSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheet(wb, LOG_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
    End If

    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("日時", "レベル", "ファイル", "行", "メッセージ")
    Set PrepareLogSheet = ws
End Function

Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then
            Set GetWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub LogMessage(ByVal wsLog As Worksheet, ByVal levelText As String, ByVal sourceFile As String, ByVal rowText As String, ByVal messageText As String)
    Dim nextRow As Long

    If wsLog Is Nothing Then Exit Sub

    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = levelText
    wsLog.Cells(nextRow, 3).Value = sourceFile
    wsLog.Cells(nextRow, 4).Value = rowText
    wsLog.Cells(nextRow, 5).Value = messageText
End Sub

Private Function BuildTargetRowMap(ByVal wsMonthly As Worksheet, ByVal lastRow As Long, ByVal wsLog As Worksheet, ByRef warnings As Long) As Object
    Dim dict As Object
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim keyText As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For rowIndex = FIRST_DATA_ROW To lastRow
        branchCode = Trim$(CStr(wsMonthly.Cells(rowIndex, COL_BRANCH_CODE).Value))
        businessCode = Trim$(CStr(wsMonthly.Cells(rowIndex, COL_BUSINESS_CODE).Value))

        If Len(branchCode) > 0 Or Len(businessCode) > 0 Then
            If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
                warnings = warnings + 1
                LogMessage wsLog, "WARN", "", CStr(rowIndex), "月次側の支店コードまたは業務コードが空のため、転記対象から外しました。"
            Else
                keyText = MakeKey(branchCode, businessCode)
                If dict.Exists(keyText) Then
                    warnings = warnings + 1
                    LogMessage wsLog, "WARN", "", CStr(rowIndex), "月次側に同じ支店コード + 業務コードが重複しています。先に見つかった行を使用します。"
                Else
                    dict.Add keyText, rowIndex
                End If
            End If
        End If
    Next rowIndex

    Set BuildTargetRowMap = dict
End Function

Private Sub ResetMonthlyValues(ByVal wsMonthly As Worksheet, ByVal lastRow As Long)
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    wsMonthly.Range(wsMonthly.Cells(FIRST_DATA_ROW, FIRST_DAY_COL), wsMonthly.Cells(lastRow, COL_MONTH_TOTAL)).ClearContents
    wsMonthly.Range(wsMonthly.Cells(FIRST_DATA_ROW, FIRST_DAY_COL), wsMonthly.Cells(lastRow, LAST_DAY_COL)).Value = 0
    wsMonthly.Range(wsMonthly.Cells(FIRST_DATA_ROW, COL_MONTH_TOTAL), wsMonthly.Cells(lastRow, COL_MONTH_TOTAL)).Value = 0
End Sub

Private Sub ProcessDailyWorkbook(ByVal fullPath As String, ByVal displayName As String, ByVal dayNumber As Long, ByVal wsMonthly As Worksheet, ByVal targetRows As Object, ByVal wsLog As Worksheet, ByRef importedRows As Long, ByRef skippedRows As Long, ByRef warnings As Long)
    Dim wbDaily As Workbook
    Dim wsDaily As Worksheet
    Dim lastDailyRow As Long
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim keyText As String
    Dim claimValue As Variant
    Dim targetRow As Long
    Dim targetCell As Range
    Dim baseValue As Double

    On Error GoTo FileError

    Set wbDaily = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set wsDaily = GetWorksheet(wbDaily, DAILY_SHEET_NAME)

    If wsDaily Is Nothing Then
        warnings = warnings + 1
        LogMessage wsLog, "ERROR", displayName, "", "日次シート '" & DAILY_SHEET_NAME & "' が見つかりません。"
        GoTo Cleanup
    End If

    lastDailyRow = LastUsedRow(wsDaily, COL_BRANCH_CODE)
    If lastDailyRow < FIRST_DATA_ROW Then
        warnings = warnings + 1
        LogMessage wsLog, "WARN", displayName, "", "日次シートにデータ行がありません。"
        GoTo Cleanup
    End If

    For rowIndex = FIRST_DATA_ROW To lastDailyRow
        branchCode = Trim$(CStr(wsDaily.Cells(rowIndex, COL_BRANCH_CODE).Value))
        businessCode = Trim$(CStr(wsDaily.Cells(rowIndex, COL_BUSINESS_CODE).Value))
        claimValue = wsDaily.Cells(rowIndex, COL_CLAIM_COUNT).Value

        If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
            skippedRows = skippedRows + 1
            warnings = warnings + 1
            LogMessage wsLog, "WARN", displayName, CStr(rowIndex), "日次側の支店コードまたは業務コードが空のため、スキップしました。"
        ElseIf IsError(claimValue) Then
            skippedRows = skippedRows + 1
            warnings = warnings + 1
            LogMessage wsLog, "WARN", displayName, CStr(rowIndex), "クレーム件数がエラー値のため、スキップしました。"
        ElseIf Len(Trim$(CStr(claimValue))) = 0 Or Not IsNumeric(claimValue) Then
            skippedRows = skippedRows + 1
            warnings = warnings + 1
            LogMessage wsLog, "WARN", displayName, CStr(rowIndex), "クレーム件数が数値ではないため、スキップしました。"
        Else
            keyText = MakeKey(branchCode, businessCode)
            If Not targetRows.Exists(keyText) Then
                skippedRows = skippedRows + 1
                warnings = warnings + 1
                LogMessage wsLog, "WARN", displayName, CStr(rowIndex), "月次側に一致する支店コード + 業務コードがないため、スキップしました: " & keyText
            Else
                targetRow = CLng(targetRows(keyText))
                Set targetCell = wsMonthly.Cells(targetRow, FIRST_DAY_COL + dayNumber - 1)

                If IsNumeric(targetCell.Value) Then
                    baseValue = CDbl(targetCell.Value)
                Else
                    baseValue = 0
                End If

                targetCell.Value = baseValue + CDbl(claimValue)
                importedRows = importedRows + 1
            End If
        End If
    Next rowIndex

    LogMessage wsLog, "INFO", displayName, "", "ファイル処理が完了しました。"

Cleanup:
    On Error Resume Next
    If Not wbDaily Is Nothing Then wbDaily.Close SaveChanges:=False
    Exit Sub

FileError:
    warnings = warnings + 1
    LogMessage wsLog, "ERROR", displayName, "", "ファイル処理中にエラーが発生しました: " & Err.Description
    Resume Cleanup
End Sub

Private Sub CalculateMonthlyTotals(ByVal wsMonthly As Worksheet, ByVal lastRow As Long)
    Dim rowIndex As Long

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For rowIndex = FIRST_DATA_ROW To lastRow
        wsMonthly.Cells(rowIndex, COL_MONTH_TOTAL).Formula = "=SUM(" & _
            wsMonthly.Range(wsMonthly.Cells(rowIndex, FIRST_DAY_COL), wsMonthly.Cells(rowIndex, LAST_DAY_COL)).Address(False, False) & ")"
    Next rowIndex

    wsMonthly.Range(wsMonthly.Cells(FIRST_DATA_ROW, FIRST_DAY_COL), wsMonthly.Cells(lastRow, COL_MONTH_TOTAL)).Calculate
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal keyColumn As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).Row
End Function

Private Function MakeKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeKey = Trim$(branchCode) & "|" & Trim$(businessCode)
End Function

Private Function JoinPath(ByVal basePath As String, ByVal childName As String) As String
    If Len(basePath) = 0 Then
        JoinPath = childName
    ElseIf Right$(basePath, 1) = Application.PathSeparator Then
        JoinPath = basePath & childName
    Else
        JoinPath = basePath & Application.PathSeparator & childName
    End If
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo NotFound
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    Exit Function

NotFound:
    FolderExists = False
End Function

Private Function ExtractDigitRun(ByVal sourceText As String, ByVal runLength As Long) As String
    Dim index As Long
    Dim currentRun As String
    Dim oneChar As String

    For index = 1 To Len(sourceText)
        oneChar = Mid$(sourceText, index, 1)
        If oneChar >= "0" And oneChar <= "9" Then
            currentRun = currentRun & oneChar
            If Len(currentRun) = runLength Then
                ExtractDigitRun = currentRun
                Exit Function
            End If
        Else
            currentRun = vbNullString
        End If
    Next index

    ExtractDigitRun = vbNullString
End Function
