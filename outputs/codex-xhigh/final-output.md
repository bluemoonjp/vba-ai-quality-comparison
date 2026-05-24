# codex-xhigh 最終出力

```vba
Option Explicit

Private Const MONTHLY_SHEET_NAME As String = "月次集計"
Private Const MASTER_SHEET_NAME As String = "マスタ"
Private Const LOG_SHEET_NAME As String = "異常系"
Private Const DAILY_SHEET_NAME As String = "日次集計"

Private Const MONTHLY_PREFIX As String = "月次クレーム集計"
Private Const DAILY_PREFIX As String = "クレーム集計"

Private Const COL_ROW_TYPE As Long = 1
Private Const COL_SUBTOTAL_GROUP As Long = 2
Private Const COL_BRANCH_CODE As Long = 4
Private Const COL_BUSINESS_CODE As Long = 6

Private Const DAY_FIRST_COL As Long = 8      ' H
Private Const DAY_LAST_COL As Long = 38      ' AL
Private Const MONTH_TOTAL_COL As Long = 39   ' AM

Public Sub ImportDailyClaimsToMonthly()
    Dim wbMonthly As Workbook
    Dim wsMonthly As Worksheet
    Dim wsMaster As Worksheet
    Dim wsLog As Worksheet

    Dim yymm As String
    Dim targetYear As Long
    Dim targetMonth As Long
    Dim daysInMonth As Long
    Dim dailyRoot As String
    Dim message As String

    Dim keyToRow As Object
    Dim branchSet As Object
    Dim businessSet As Object
    Dim monthlyBranchSet As Object
    Dim monthlyBusinessSet As Object
    Dim dateFiles As Object

    Dim detailRows As Collection
    Dim subtotalRows As Collection
    Dim sortedDates As Collection
    Dim sortedPaths As Collection
    Dim paths As Collection

    Dim fso As Object
    Dim dateKey As Variant
    Dim selectedPath As String
    Dim duplicateIndex As Long

    Dim logRow As Long
    Dim candidateFileCount As Long
    Dim processedFileCount As Long
    Dim importedRowCount As Long
    Dim importedAmount As Double
    Dim issueCount As Long

    Dim branchFallbackNeeded As Boolean
    Dim businessFallbackNeeded As Boolean

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim settingsCaptured As Boolean
    Dim fatalMessage As String

    On Error GoTo FatalError

    Set wbMonthly = ThisWorkbook
    Set wsLog = GetOrCreateWorksheet(wbMonthly, LOG_SHEET_NAME)
    PrepareLogSheet wsLog, logRow

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wsMonthly = GetRequiredWorksheet(wbMonthly, MONTHLY_SHEET_NAME)

    If Not TryParseMonthlyYYMM(wbMonthly.Name, yymm, targetYear, targetMonth, message) Then
        AddLog wsLog, logRow, "ERROR", "MONTHLY_FILENAME_ERROR", wbMonthly.FullName, "", 0, "", "", "", "", message, "処理中止"
        Err.Raise vbObjectError + 1001, , message
    End If

    If Len(wbMonthly.Path) = 0 Then
        message = "月次ブックが保存されていないため、dailyフォルダを判定できません。"
        AddLog wsLog, logRow, "ERROR", "WORKBOOK_NOT_SAVED", wbMonthly.FullName, "", 0, "", "", "", "", message, "処理中止"
        Err.Raise vbObjectError + 1002, , message
    End If

    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))
    dailyRoot = wbMonthly.Path & Application.PathSeparator & "daily" & Application.PathSeparator & yymm

    Set keyToRow = CreateTextDictionary()
    Set branchSet = CreateTextDictionary()
    Set businessSet = CreateTextDictionary()
    Set monthlyBranchSet = CreateTextDictionary()
    Set monthlyBusinessSet = CreateTextDictionary()
    Set dateFiles = CreateTextDictionary()

    Set detailRows = New Collection
    Set subtotalRows = New Collection

    BuildMonthlyIndex wsMonthly, keyToRow, detailRows, subtotalRows, monthlyBranchSet, monthlyBusinessSet, wsLog, logRow

    Set wsMaster = GetWorksheetIfExists(wbMonthly, MASTER_SHEET_NAME)
    If wsMaster Is Nothing Then
        branchFallbackNeeded = True
        businessFallbackNeeded = True
        AddLog wsLog, logRow, "WARN", "MASTER_SHEET_MISSING", wbMonthly.FullName, MASTER_SHEET_NAME, 0, "", "", "", "", _
               "マスタシートがないため、月次集計のDETAIL行からコード一覧を補完します。", "継続"
    Else
        LoadMasterCodeSets wsMaster, branchSet, businessSet, wsLog, logRow, branchFallbackNeeded, businessFallbackNeeded
    End If

    If branchFallbackNeeded Or branchSet.Count = 0 Then
        CopyDictionaryKeys monthlyBranchSet, branchSet
        AddLog wsLog, logRow, "WARN", "MASTER_BRANCH_FALLBACK", wbMonthly.FullName, MONTHLY_SHEET_NAME, 0, "", "", "", "", _
               "支店コード一覧を月次集計のDETAIL行から補完しました。", "継続"
    End If

    If businessFallbackNeeded Or businessSet.Count = 0 Then
        CopyDictionaryKeys monthlyBusinessSet, businessSet
        AddLog wsLog, logRow, "WARN", "MASTER_BUSINESS_FALLBACK", wbMonthly.FullName, MONTHLY_SHEET_NAME, 0, "", "", "", "", _
               "業務コード一覧を月次集計のDETAIL行から補完しました。", "継続"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dailyRoot) Then
        message = "日次ファイル探索フォルダが存在しません: " & ToLogPath(dailyRoot)
        AddLog wsLog, logRow, "ERROR", "DAILY_ROOT_MISSING", dailyRoot, "", 0, "", "", "", "", message, "処理中止"
        Err.Raise vbObjectError + 1003, , message
    End If

    CollectDailyFiles fso.GetFolder(dailyRoot), yymm, dateFiles, wsLog, logRow, candidateFileCount

    ClearTransferArea wsMonthly, detailRows

    If dateFiles.Count = 0 Then
        AddLog wsLog, logRow, "WARN", "NO_DAILY_FILES", dailyRoot, "", 0, "", "", "", "", _
               "処理対象となる日次ファイルがありませんでした。", "月次集計のDETAIL日別欄はクリア済み"
    End If

    Set sortedDates = SortDictionaryKeys(dateFiles)

    For Each dateKey In sortedDates
        Set paths = dateFiles(CStr(dateKey))
        Set sortedPaths = SortTextCollection(paths)
        selectedPath = CStr(sortedPaths(1))

        If sortedPaths.Count > 1 Then
            For duplicateIndex = 2 To sortedPaths.Count
                AddLog wsLog, logRow, "WARN", "DUPLICATE_DAILY_FILE_SKIPPED", CStr(sortedPaths(duplicateIndex)), "", 0, _
                       FormatDateKey(CStr(dateKey)), "", "", "", _
                       "同じ日付の日次ファイルが複数あります。フルパス昇順の先頭1件を採用します。採用ファイル=" & ToLogPath(selectedPath), _
                       "重複ファイルをスキップ"
            Next duplicateIndex
        End If

        ProcessDailyWorkbook selectedPath, CStr(dateKey), keyToRow, branchSet, businessSet, wsMonthly, wsLog, logRow, _
                             processedFileCount, importedRowCount, importedAmount
    Next dateKey

    RecalculateMonthlyTotals wsMonthly, detailRows, subtotalRows, daysInMonth, wsLog, logRow

    issueCount = logRow - 2
    AddLog wsLog, logRow, "INFO", "SUMMARY", wbMonthly.FullName, "", 0, "", "", "", CStr(importedAmount), _
           "候補ファイル=" & candidateFileCount & _
           ", 取込ファイル=" & processedFileCount & _
           ", 転記行=" & importedRowCount & _
           ", 転記件数合計=" & importedAmount & _
           ", ログ件数=" & issueCount, "完了"

    wsLog.Columns("A:L").AutoFit

    MsgBox "月次クレーム集計が完了しました。" & vbCrLf & _
           "対象月: " & yymm & vbCrLf & _
           "取込ファイル: " & processedFileCount & vbCrLf & _
           "転記行: " & importedRowCount & vbCrLf & _
           "ログ件数: " & issueCount & vbCrLf & _
           "詳細は「" & LOG_SHEET_NAME & "」シートを確認してください。", vbInformation

CleanExit:
    On Error Resume Next
    If settingsCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.DisplayAlerts = oldDisplayAlerts
        Application.Calculation = oldCalculation
    End If
    On Error GoTo 0
    Exit Sub

FatalError:
    fatalMessage = Err.Description
    On Error Resume Next
    If Not wsLog Is Nothing Then
        AddLog wsLog, logRow, "ERROR", "FATAL_ERROR", wbMonthly.FullName, "", 0, "", "", "", "", fatalMessage, "処理中止"
        wsLog.Columns("A:L").AutoFit
    End If
    MsgBox "処理を中止しました。" & vbCrLf & SanitizeLogText(fatalMessage) & vbCrLf & _
           "詳細は「" & LOG_SHEET_NAME & "」シートを確認してください。", vbExclamation
    On Error GoTo 0
    Resume CleanExit
End Sub

Private Sub BuildMonthlyIndex(ByVal wsMonthly As Worksheet, ByVal keyToRow As Object, ByVal detailRows As Collection, _
                              ByVal subtotalRows As Collection, ByVal monthlyBranchSet As Object, _
                              ByVal monthlyBusinessSet As Object, ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim rowType As String
    Dim branchCode As String
    Dim businessCode As String
    Dim key As String

    lastRow = GetLastRowInColumns(wsMonthly, 1, MONTH_TOTAL_COL)

    For rowNumber = 2 To lastRow
        rowType = UCase$(NormalizeValue(wsMonthly.Cells(rowNumber, COL_ROW_TYPE).Value2))

        If rowType = "DETAIL" Then
            detailRows.Add rowNumber

            branchCode = CodeValue(wsMonthly.Cells(rowNumber, COL_BRANCH_CODE))
            businessCode = CodeValue(wsMonthly.Cells(rowNumber, COL_BUSINESS_CODE))

            AddKeyIfMissing monthlyBranchSet, branchCode
            AddKeyIfMissing monthlyBusinessSet, businessCode

            If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
                AddLog wsLog, logRow, "ERROR", "MONTHLY_DETAIL_KEY_BLANK", ThisWorkbook.FullName, MONTHLY_SHEET_NAME, rowNumber, "", branchCode, businessCode, "", _
                       "月次集計のDETAIL行で支店コードまたは業務コードが空欄です。", "この行は転記キー対象外"
            Else
                key = MakeMonthlyKey(branchCode, businessCode)
                If keyToRow.Exists(key) Then
                    AddLog wsLog, logRow, "ERROR", "MONTHLY_KEY_DUPLICATE", ThisWorkbook.FullName, MONTHLY_SHEET_NAME, rowNumber, "", branchCode, businessCode, "", _
                           "月次集計に同じ支店コード+業務コードのDETAIL行が複数あります。先に見つかった行へ転記します。", "継続"
                Else
                    keyToRow.Add key, rowNumber
                End If
            End If

        ElseIf rowType = "SUBTOTAL" Then
            subtotalRows.Add rowNumber
        End If
    Next rowNumber
End Sub

Private Sub ClearTransferArea(ByVal wsMonthly As Worksheet, ByVal detailRows As Collection)
    Dim item As Variant
    Dim rowNumber As Long

    For Each item In detailRows
        rowNumber = CLng(item)
        wsMonthly.Range(wsMonthly.Cells(rowNumber, DAY_FIRST_COL), wsMonthly.Cells(rowNumber, DAY_LAST_COL)).ClearContents
    Next item
End Sub

Private Sub CollectDailyFiles(ByVal folder As Object, ByVal expectedYYMM As String, ByVal dateFiles As Object, _
                              ByVal wsLog As Worksheet, ByRef logRow As Long, ByRef candidateFileCount As Long)
    Dim fileItem As Object
    Dim subFolder As Object
    Dim dayNumber As Long
    Dim fileDate As Date
    Dim dateKey As String
    Dim category As String
    Dim message As String

    For Each fileItem In folder.Files
        If ShouldEvaluateDailyFile(fileItem.Name) Then
            candidateFileCount = candidateFileCount + 1

            If TryParseDailyFileName(fileItem.Name, expectedYYMM, dayNumber, fileDate, category, message) Then
                dateKey = Format$(fileDate, "yyyymmdd")
                AddFileToDateDictionary dateFiles, dateKey, fileItem.Path
            Else
                AddLog wsLog, logRow, "WARN", category, fileItem.Path, "", 0, "", "", "", "", message, "スキップ"
            End If
        End If
    Next fileItem

    For Each subFolder In folder.SubFolders
        CollectDailyFiles subFolder, expectedYYMM, dateFiles, wsLog, logRow, candidateFileCount
    Next subFolder
End Sub

Private Sub ProcessDailyWorkbook(ByVal filePath As String, ByVal dateKey As String, ByVal keyToRow As Object, _
                                 ByVal branchSet As Object, ByVal businessSet As Object, ByVal wsMonthly As Worksheet, _
                                 ByVal wsLog As Worksheet, ByRef logRow As Long, ByRef processedFileCount As Long, _
                                 ByRef importedRowCount As Long, ByRef importedAmount As Double)
    Dim wbDaily As Workbook
    Dim wsDaily As Worksheet
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim dayNumber As Long
    Dim displayDate As String

    Dim statusValue As String
    Dim branchCode As String
    Dim businessCode As String
    Dim claimRaw As Variant
    Dim claimText As String
    Dim claimCount As Double
    Dim validRow As Boolean
    Dim key As String
    Dim targetRow As Long

    dayNumber = CLng(Right$(dateKey, 2))
    displayDate = FormatDateKey(dateKey)

    On Error GoTo OpenError
    Set wbDaily = Workbooks.Open(Filename:=filePath, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)

    On Error GoTo ProcessError

    If Not WorksheetExists(wbDaily, DAILY_SHEET_NAME) Then
        AddLog wsLog, logRow, "ERROR", "DAILY_SHEET_MISSING", filePath, DAILY_SHEET_NAME, 0, displayDate, "", "", "", _
               "日次ブックに「" & DAILY_SHEET_NAME & "」シートがありません。", "ファイルをスキップ"
        GoTo Cleanup
    End If

    Set wsDaily = wbDaily.Worksheets(DAILY_SHEET_NAME)
    ValidateDailyHeaders wsDaily, filePath, wsLog, logRow

    processedFileCount = processedFileCount + 1
    lastRow = GetLastRowInColumns(wsDaily, 1, 7)

    For rowNumber = 2 To lastRow
        If IsDailyRowCompletelyBlank(wsDaily, rowNumber) Then GoTo NextRow

        statusValue = LCase$(NormalizeValue(wsDaily.Cells(rowNumber, 1).Value2))
        branchCode = CodeValue(wsDaily.Cells(rowNumber, 2))
        businessCode = CodeValue(wsDaily.Cells(rowNumber, 4))
        claimRaw = wsDaily.Cells(rowNumber, 6).Value
        claimText = SafeCellText(wsDaily.Cells(rowNumber, 6))

        If statusValue <> "ok" Then
            AddLog wsLog, logRow, "WARN", "NON_OK_STATUS", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, _
                   "処理区分がokではありません。処理区分=" & NormalizeValue(wsDaily.Cells(rowNumber, 1).Value2), "転記しない"
            GoTo NextRow
        End If

        validRow = True

        If Not branchSet.Exists(branchCode) Then
            AddLog wsLog, logRow, "ERROR", "UNKNOWN_BRANCH", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, _
                   "マスタに存在しない支店コードです。", "転記しない"
            validRow = False
        End If

        If Not businessSet.Exists(businessCode) Then
            AddLog wsLog, logRow, "ERROR", "UNKNOWN_BUSINESS", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, _
                   "マスタに存在しない業務コードです。", "転記しない"
            validRow = False
        End If

        If IsBlankValue(claimRaw) Then
            AddLog wsLog, logRow, "ERROR", "EMPTY_CLAIM_COUNT", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, _
                   "クレーム件数が空欄です。", "転記しない"
            validRow = False
        ElseIf IsError(claimRaw) Or Not IsNumeric(claimRaw) Then
            AddLog wsLog, logRow, "ERROR", "NON_NUMERIC_CLAIM_COUNT", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, _
                   "クレーム件数が数値ではありません。", "転記しない"
            validRow = False
        Else
            claimCount = CDbl(claimRaw)
        End If

        If validRow Then
            key = MakeMonthlyKey(branchCode, businessCode)

            If Not keyToRow.Exists(key) Then
                AddLog wsLog, logRow, "ERROR", "MONTHLY_KEY_MISSING", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, _
                       "月次集計に該当する支店コード+業務コードのDETAIL行がありません。", "転記しない"
            Else
                targetRow = CLng(keyToRow(key))
                AddClaimToMonthlyCell wsMonthly, targetRow, dayNumber, claimCount
                importedRowCount = importedRowCount + 1
                importedAmount = importedAmount + claimCount
            End If
        End If

NextRow:
    Next rowNumber

Cleanup:
    On Error Resume Next
    If Not wbDaily Is Nothing Then wbDaily.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

OpenError:
    AddLog wsLog, logRow, "ERROR", "DAILY_FILE_OPEN_ERROR", filePath, "", 0, displayDate, "", "", "", Err.Description, "ファイルをスキップ"
    Resume Cleanup

ProcessError:
    AddLog wsLog, logRow, "ERROR", "DAILY_FILE_PROCESS_ERROR", filePath, DAILY_SHEET_NAME, rowNumber, displayDate, branchCode, businessCode, claimText, Err.Description, "ファイル処理を中断"
    Resume Cleanup
End Sub

Private Sub RecalculateMonthlyTotals(ByVal wsMonthly As Worksheet, ByVal detailRows As Collection, ByVal subtotalRows As Collection, _
                                     ByVal daysInMonth As Long, ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim groupDayTotals As Object
    Dim groupKnown As Object
    Dim item As Variant
    Dim rowNumber As Long
    Dim dayNumber As Long
    Dim colNumber As Long
    Dim rowTotal As Double
    Dim groupName As String
    Dim key As String
    Dim valueNumber As Double
    Dim subtotalValue As Double

    Set groupDayTotals = CreateTextDictionary()
    Set groupKnown = CreateTextDictionary()

    For Each item In detailRows
        rowNumber = CLng(item)
        groupName = NormalizeValue(wsMonthly.Cells(rowNumber, COL_SUBTOTAL_GROUP).Value2)
        rowTotal = 0

        If Len(groupName) > 0 Then AddKeyIfMissing groupKnown, groupName

        For dayNumber = 1 To 31
            colNumber = DAY_FIRST_COL + dayNumber - 1

            If dayNumber <= daysInMonth Then
                valueNumber = NumericCellValue(wsMonthly.Cells(rowNumber, colNumber))

                If IsEffectivelyZero(valueNumber) Then
                    wsMonthly.Cells(rowNumber, colNumber).ClearContents
                Else
                    rowTotal = rowTotal + valueNumber

                    If Len(groupName) > 0 Then
                        key = MakeGroupDayKey(groupName, dayNumber)
                        If groupDayTotals.Exists(key) Then
                            groupDayTotals(key) = CDbl(groupDayTotals(key)) + valueNumber
                        Else
                            groupDayTotals.Add key, valueNumber
                        End If
                    End If
                End If
            Else
                wsMonthly.Cells(rowNumber, colNumber).ClearContents
            End If
        Next dayNumber

        SetCellNumberOrBlank wsMonthly.Cells(rowNumber, MONTH_TOTAL_COL), rowTotal
    Next item

    For Each item In subtotalRows
        rowNumber = CLng(item)
        groupName = NormalizeValue(wsMonthly.Cells(rowNumber, COL_SUBTOTAL_GROUP).Value2)
        rowTotal = 0

        If Len(groupName) = 0 Then
            AddLog wsLog, logRow, "WARN", "SUBTOTAL_GROUP_BLANK", ThisWorkbook.FullName, MONTHLY_SHEET_NAME, rowNumber, "", "", "", "", _
                   "SUBTOTAL行の小計グループが空欄です。", "小計は空欄として再計算"
        ElseIf Not groupKnown.Exists(groupName) Then
            AddLog wsLog, logRow, "WARN", "SUBTOTAL_GROUP_WITHOUT_DETAIL", ThisWorkbook.FullName, MONTHLY_SHEET_NAME, rowNumber, "", "", "", "", _
                   "同じ小計グループのDETAIL行がありません。小計グループ=" & groupName, "小計は空欄として再計算"
        End If

        For dayNumber = 1 To 31
            colNumber = DAY_FIRST_COL + dayNumber - 1

            If dayNumber <= daysInMonth Then
                key = MakeGroupDayKey(groupName, dayNumber)
                If groupDayTotals.Exists(key) Then
                    subtotalValue = CDbl(groupDayTotals(key))
                Else
                    subtotalValue = 0
                End If

                SetCellNumberOrBlank wsMonthly.Cells(rowNumber, colNumber), subtotalValue
                rowTotal = rowTotal + subtotalValue
            Else
                wsMonthly.Cells(rowNumber, colNumber).ClearContents
            End If
        Next dayNumber

        SetCellNumberOrBlank wsMonthly.Cells(rowNumber, MONTH_TOTAL_COL), rowTotal
    Next item
End Sub

Private Sub LoadMasterCodeSets(ByVal wsMaster As Worksheet, ByVal branchSet As Object, ByVal businessSet As Object, _
                               ByVal wsLog As Worksheet, ByRef logRow As Long, _
                               ByRef branchFallbackNeeded As Boolean, ByRef businessFallbackNeeded As Boolean)
    Dim branchMasterFound As Boolean
    Dim businessMasterFound As Boolean
    Dim relationBranchFound As Boolean
    Dim relationBusinessFound As Boolean
    Dim missingMessage As String

    branchMasterFound = CollectCodesFromHeaderRow(wsMaster, 2, "支店コード", 3, 11, branchSet)
    If Not branchMasterFound Then
        AddLog wsLog, logRow, "WARN", "MASTER_BRANCH_HEADER_MISSING", ThisWorkbook.FullName, MASTER_SHEET_NAME, 2, "", "", "", "", _
               "支店マスタ見出し行で「支店コード」が見つかりません。", "月次DETAILから補完予定"
    End If

    businessMasterFound = CollectCodesFromHeaderRow(wsMaster, 12, "業務コード", 13, 23, businessSet)
    If Not businessMasterFound Then
        AddLog wsLog, logRow, "WARN", "MASTER_BUSINESS_HEADER_MISSING", ThisWorkbook.FullName, MASTER_SHEET_NAME, 12, "", "", "", "", _
               "業務マスタ見出し行で「業務コード」が見つかりません。", "月次DETAILから補完予定"
    End If

    relationBranchFound = CollectCodesFromHeaderRow(wsMaster, 24, "支店コード", 25, 0, branchSet)
    relationBusinessFound = CollectCodesFromHeaderRow(wsMaster, 24, "業務コード", 25, 0, businessSet)

    If Not relationBranchFound Or Not relationBusinessFound Then
        missingMessage = "支店・業務対応表見出し行で不足があります。"
        If Not relationBranchFound Then missingMessage = missingMessage & " 支店コードなし。"
        If Not relationBusinessFound Then missingMessage = missingMessage & " 業務コードなし。"

        AddLog wsLog, logRow, "WARN", "MASTER_RELATION_HEADER_MISSING", ThisWorkbook.FullName, MASTER_SHEET_NAME, 24, "", "", "", "", _
               missingMessage, "不足分は月次DETAILから補完予定"
    End If

    branchFallbackNeeded = ((Not branchMasterFound) And (Not relationBranchFound)) Or branchSet.Count = 0
    businessFallbackNeeded = ((Not businessMasterFound) And (Not relationBusinessFound)) Or businessSet.Count = 0
End Sub

Private Function CollectCodesFromHeaderRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String, _
                                           ByVal startRow As Long, ByVal endRow As Long, ByVal codeSet As Object) As Boolean
    Dim headerCol As Long
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim codeValue As String

    headerCol = FindHeaderColumn(ws, headerRow, headerText)
    If headerCol = 0 Then Exit Function

    CollectCodesFromHeaderRow = True

    If endRow > 0 Then
        lastRow = endRow
    Else
        lastRow = ws.Cells(ws.Rows.Count, headerCol).End(xlUp).Row
    End If

    If lastRow < startRow Then Exit Function

    For rowNumber = startRow To lastRow
        codeValue = CodeValue(ws.Cells(rowNumber, headerCol))
        If Len(codeValue) > 0 And codeValue <> headerText Then
            AddKeyIfMissing codeSet, codeValue
        End If
    Next rowNumber
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim colNumber As Long

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    If lastCol = 1 And Len(NormalizeValue(ws.Cells(headerRow, 1).Value2)) = 0 Then Exit Function

    For colNumber = 1 To lastCol
        If NormalizeValue(ws.Cells(headerRow, colNumber).Value2) = headerText Then
            FindHeaderColumn = colNumber
            Exit Function
        End If
    Next colNumber
End Function

Private Sub ValidateDailyHeaders(ByVal wsDaily As Worksheet, ByVal filePath As String, ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim expectedHeaders As Variant
    Dim i As Long
    Dim actualValue As String
    Dim mismatch As String

    expectedHeaders = Array("処理区分", "支店コード", "支店名", "業務コード", "業務名", "クレーム件数", "備考")

    For i = 0 To UBound(expectedHeaders)
        actualValue = NormalizeValue(wsDaily.Cells(1, i + 1).Value2)
        If actualValue <> CStr(expectedHeaders(i)) Then
            mismatch = mismatch & "列" & ColumnLetter(i + 1) & ": expected=" & CStr(expectedHeaders(i)) & ", actual=" & actualValue & "; "
        End If
    Next i

    If Len(mismatch) > 0 Then
        AddLog wsLog, logRow, "WARN", "DAILY_HEADER_MISMATCH", filePath, DAILY_SHEET_NAME, 1, "", "", "", "", _
               "日次集計シートの見出しが想定と異なります。固定列位置で処理を継続します。 " & mismatch, "継続"
    End If
End Sub

Private Sub AddClaimToMonthlyCell(ByVal wsMonthly As Worksheet, ByVal rowNumber As Long, ByVal dayNumber As Long, ByVal amount As Double)
    Dim targetCell As Range
    Dim currentValue As Variant

    If IsEffectivelyZero(amount) Then Exit Sub

    Set targetCell = wsMonthly.Cells(rowNumber, DAY_FIRST_COL + dayNumber - 1)
    currentValue = targetCell.Value

    If Not IsError(currentValue) And IsNumeric(currentValue) And Len(Trim$(CStr(currentValue))) > 0 Then
        targetCell.Value = CDbl(currentValue) + amount
    Else
        targetCell.Value = amount
    End If
End Sub

Private Function TryParseMonthlyYYMM(ByVal fileName As String, ByRef yymm As String, ByRef targetYear As Long, _
                                     ByRef targetMonth As Long, ByRef message As String) As Boolean
    Dim baseName As String
    Dim datePart As String

    If LCase$(FileExtension(fileName)) <> "xlsm" Then
        message = "月次ブックの拡張子が.xlsmではありません。ファイル名=" & fileName
        Exit Function
    End If

    baseName = FileBaseName(fileName)

    If Left$(baseName, Len(MONTHLY_PREFIX)) <> MONTHLY_PREFIX Then
        message = "月次ブック名が「" & MONTHLY_PREFIX & "YYMM.xlsm」形式ではありません。ファイル名=" & fileName
        Exit Function
    End If

    datePart = Mid$(baseName, Len(MONTHLY_PREFIX) + 1)

    If Len(datePart) <> 4 Or Not IsAllDigits(datePart) Then
        message = "月次ブック名のYYMM部分が4桁数字ではありません。ファイル名=" & fileName
        Exit Function
    End If

    targetYear = 2000 + CLng(Left$(datePart, 2))
    targetMonth = CLng(Right$(datePart, 2))

    If targetMonth < 1 Or targetMonth > 12 Then
        message = "月次ブック名の月が不正です。YYMM=" & datePart
        Exit Function
    End If

    yymm = datePart
    TryParseMonthlyYYMM = True
End Function

Private Function TryParseDailyFileName(ByVal fileName As String, ByVal expectedYYMM As String, ByRef dayNumber As Long, _
                                       ByRef fileDate As Date, ByRef logCategory As String, ByRef logMessage As String) As Boolean
    Dim baseName As String
    Dim datePart As String
    Dim yy As Long
    Dim mm As Long
    Dim dd As Long
    Dim maxDay As Long

    If LCase$(FileExtension(fileName)) <> "xlsx" Then
        logCategory = "FILENAME_FORMAT"
        logMessage = "日次ファイル名が「" & DAILY_PREFIX & "YYMMDD.xlsx」形式ではありません。ファイル名=" & fileName
        Exit Function
    End If

    baseName = FileBaseName(fileName)

    If Left$(baseName, Len(DAILY_PREFIX)) <> DAILY_PREFIX Then
        logCategory = "FILENAME_FORMAT"
        logMessage = "日次ファイル名が「" & DAILY_PREFIX & "YYMMDD.xlsx」形式ではありません。ファイル名=" & fileName
        Exit Function
    End If

    datePart = Mid$(baseName, Len(DAILY_PREFIX) + 1)

    If Len(datePart) <> 6 Or Not IsAllDigits(datePart) Then
        logCategory = "FILENAME_FORMAT"
        logMessage = "日次ファイル名のYYMMDD部分が6桁数字ではありません。ファイル名=" & fileName
        Exit Function
    End If

    If Left$(datePart, 4) <> expectedYYMM Then
        logCategory = "TARGET_MONTH_MISMATCH"
        logMessage = "日次ファイルの対象月が月次ブックと異なります。期待YYMM=" & expectedYYMM & ", 日次YYMM=" & Left$(datePart, 4)
        Exit Function
    End If

    yy = CLng(Left$(datePart, 2))
    mm = CLng(Mid$(datePart, 3, 2))
    dd = CLng(Right$(datePart, 2))

    If mm < 1 Or mm > 12 Then
        logCategory = "FILENAME_FORMAT"
        logMessage = "日次ファイル名の月が不正です。ファイル名=" & fileName
        Exit Function
    End If

    maxDay = Day(DateSerial(2000 + yy, mm + 1, 0))
    If dd < 1 Or dd > maxDay Then
        logCategory = "FILENAME_FORMAT"
        logMessage = "日次ファイル名の日が不正です。ファイル名=" & fileName
        Exit Function
    End If

    dayNumber = dd
    fileDate = DateSerial(2000 + yy, mm, dd)
    TryParseDailyFileName = True
End Function

Private Sub PrepareLogSheet(ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim headers As Variant
    Dim i As Long

    wsLog.Cells.ClearContents

    headers = Array("実行時刻", "レベル", "カテゴリ", "ファイル", "シート", "行", "日付", "支店コード", "業務コード", "クレーム件数", "メッセージ", "処理")

    For i = 0 To UBound(headers)
        wsLog.Cells(1, i + 1).Value = headers(i)
    Next i

    logRow = 2
End Sub

Private Sub AddLog(ByVal wsLog As Worksheet, ByRef logRow As Long, ByVal level As String, ByVal category As String, _
                   ByVal filePath As String, ByVal sheetName As String, ByVal rowNumber As Long, ByVal dateText As String, _
                   ByVal branchCode As String, ByVal businessCode As String, ByVal claimValue As String, _
                   ByVal message As String, ByVal actionText As String)
    With wsLog
        .Cells(logRow, 1).Value = Now
        .Cells(logRow, 2).Value = level
        .Cells(logRow, 3).Value = category
        .Cells(logRow, 4).Value = ToLogPath(filePath)
        .Cells(logRow, 5).Value = sheetName
        If rowNumber > 0 Then .Cells(logRow, 6).Value = rowNumber
        .Cells(logRow, 7).Value = dateText
        .Cells(logRow, 8).Value = branchCode
        .Cells(logRow, 9).Value = businessCode
        .Cells(logRow, 10).Value = claimValue
        .Cells(logRow, 11).Value = SanitizeLogText(message)
        .Cells(logRow, 12).Value = SanitizeLogText(actionText)
    End With

    logRow = logRow + 1
End Sub

Private Function CreateTextDictionary() As Object
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set CreateTextDictionary = dict
End Function

Private Sub AddKeyIfMissing(ByVal dict As Object, ByVal key As String)
    If Len(key) = 0 Then Exit Sub
    If Not dict.Exists(key) Then dict.Add key, True
End Sub

Private Sub CopyDictionaryKeys(ByVal sourceDict As Object, ByVal targetDict As Object)
    Dim key As Variant

    For Each key In sourceDict.Keys
        AddKeyIfMissing targetDict, CStr(key)
    Next key
End Sub

Private Sub AddFileToDateDictionary(ByVal dateFiles As Object, ByVal dateKey As String, ByVal filePath As String)
    Dim paths As Collection

    If dateFiles.Exists(dateKey) Then
        Set paths = dateFiles(dateKey)
    Else
        Set paths = New Collection
        dateFiles.Add dateKey, paths
    End If

    paths.Add filePath
End Sub

Private Function SortDictionaryKeys(ByVal dict As Object) As Collection
    Dim keys As Collection
    Dim key As Variant

    Set keys = New Collection

    For Each key In dict.Keys
        keys.Add CStr(key)
    Next key

    Set SortDictionaryKeys = SortTextCollection(keys)
End Function

Private Function SortTextCollection(ByVal source As Collection) As Collection
    Dim result As Collection
    Dim arr() As String
    Dim i As Long

    Set result = New Collection

    If source.Count = 0 Then
        Set SortTextCollection = result
        Exit Function
    End If

    ReDim arr(1 To source.Count)

    For i = 1 To source.Count
        arr(i) = CStr(source(i))
    Next i

    If UBound(arr) > LBound(arr) Then
        QuickSortStrings arr, LBound(arr), UBound(arr)
    End If

    For i = LBound(arr) To UBound(arr)
        result.Add arr(i)
    Next i

    Set SortTextCollection = result
End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim low As Long
    Dim high As Long
    Dim pivot As String
    Dim temp As String

    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high
        Do While StrComp(arr(low), pivot, vbTextCompare) < 0
            low = low + 1
        Loop

        Do While StrComp(arr(high), pivot, vbTextCompare) > 0
            high = high - 1
        Loop

        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortStrings arr, first, high
    If low < last Then QuickSortStrings arr, low, last
End Sub

Private Function MakeMonthlyKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeMonthlyKey = branchCode & "|" & businessCode
End Function

Private Function MakeGroupDayKey(ByVal groupName As String, ByVal dayNumber As Long) As String
    MakeGroupDayKey = groupName & "|" & CStr(dayNumber)
End Function

Private Function ShouldEvaluateDailyFile(ByVal fileName As String) As Boolean
    If Left$(fileName, 2) = "~$" Then Exit Function

    If LCase$(FileExtension(fileName)) = "xlsx" Then
        ShouldEvaluateDailyFile = True
    ElseIf Left$(fileName, Len(DAILY_PREFIX)) = DAILY_PREFIX Then
        ShouldEvaluateDailyFile = True
    End If
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not ws Is Nothing
End Function

Private Function GetWorksheetIfExists(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetIfExists = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheetIfExists(wb, sheetName)

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set GetOrCreateWorksheet = ws
End Function

Private Function GetRequiredWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Set GetRequiredWorksheet = GetWorksheetIfExists(wb, sheetName)

    If GetRequiredWorksheet Is Nothing Then
        Err.Raise vbObjectError + 2001, , "必要なシートがありません: " & sheetName
    End If
End Function

Private Function GetLastRowInColumns(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim colNumber As Long
    Dim rowNumber As Long
    Dim maxRow As Long

    For colNumber = firstCol To lastCol
        rowNumber = ws.Cells(ws.Rows.Count, colNumber).End(xlUp).Row
        If rowNumber > maxRow Then maxRow = rowNumber
    Next colNumber

    If maxRow < 1 Then maxRow = 1
    GetLastRowInColumns = maxRow
End Function

Private Function IsDailyRowCompletelyBlank(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    Dim colNumber As Long

    For colNumber = 1 To 7
        If Len(NormalizeValue(ws.Cells(rowNumber, colNumber).Value2)) > 0 Then
            IsDailyRowCompletelyBlank = False
            Exit Function
        End If
    Next colNumber

    IsDailyRowCompletelyBlank = True
End Function

Private Function IsBlankValue(ByVal value As Variant) As Boolean
    If IsError(value) Then
        IsBlankValue = False
    ElseIf IsEmpty(value) Then
        IsBlankValue = True
    ElseIf Trim$(CStr(value)) = "" Then
        IsBlankValue = True
    End If
End Function

Private Function NumericCellValue(ByVal cell As Range) As Double
    Dim value As Variant

    value = cell.Value

    If Not IsError(value) Then
        If IsNumeric(value) And Len(Trim$(CStr(value))) > 0 Then
            NumericCellValue = CDbl(value)
        End If
    End If
End Function

Private Sub SetCellNumberOrBlank(ByVal cell As Range, ByVal valueNumber As Double)
    If IsEffectivelyZero(valueNumber) Then
        cell.ClearContents
    Else
        cell.Value = valueNumber
    End If
End Sub

Private Function IsEffectivelyZero(ByVal valueNumber As Double) As Boolean
    IsEffectivelyZero = (Abs(valueNumber) < 0.0000001)
End Function

Private Function NormalizeValue(ByVal value As Variant) As String
    If IsError(value) Then
        NormalizeValue = "#ERROR"
    Else
        NormalizeValue = Trim$(CStr(value))
    End If
End Function

Private Function CodeValue(ByVal cell As Range) As String
    Dim textValue As String

    On Error Resume Next
    textValue = Trim$(CStr(cell.Text))
    On Error GoTo 0

    If Len(textValue) > 0 And InStr(1, textValue, "#", vbBinaryCompare) = 0 Then
        CodeValue = textValue
    Else
        CodeValue = NormalizeValue(cell.Value2)
    End If
End Function

Private Function SafeCellText(ByVal cell As Range) As String
    On Error GoTo Fallback
    SafeCellText = Trim$(CStr(cell.Text))
    Exit Function

Fallback:
    SafeCellText = NormalizeValue(cell.Value2)
End Function

Private Function IsAllDigits(ByVal textValue As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    IsAllDigits = True
End Function

Private Function FileExtension(ByVal fileName As String) As String
    Dim pos As Long

    pos = InStrRev(fileName, ".")
    If pos > 0 Then FileExtension = Mid$(fileName, pos + 1)
End Function

Private Function FileBaseName(ByVal fileName As String) As String
    Dim pos As Long

    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        FileBaseName = Left$(fileName, pos - 1)
    Else
        FileBaseName = fileName
    End If
End Function

Private Function FormatDateKey(ByVal dateKey As String) As String
    FormatDateKey = Left$(dateKey, 4) & "-" & Mid$(dateKey, 5, 2) & "-" & Right$(dateKey, 2)
End Function

Private Function ColumnLetter(ByVal colNumber As Long) As String
    Dim dividend As Long
    Dim modulo As Long
    Dim result As String

    dividend = colNumber

    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        result = Chr$(65 + modulo) & result
        dividend = (dividend - modulo - 1) \ 26
    Loop

    ColumnLetter = result
End Function

Private Function ToLogPath(ByVal pathValue As String) As String
    Dim basePath As String
    Dim normalizedPath As String
    Dim normalizedBase As String
    Dim baseWithSep As String
    Dim sep As String

    If Len(pathValue) = 0 Then Exit Function

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        ToLogPath = pathValue
        Exit Function
    End If

    sep = Application.PathSeparator
    normalizedPath = Replace(pathValue, "/", sep)
    normalizedBase = Replace(basePath, "/", sep)

    If StrComp(normalizedPath, ThisWorkbook.FullName, vbTextCompare) = 0 Then
        ToLogPath = ThisWorkbook.Name
        Exit Function
    End If

    If StrComp(normalizedPath, normalizedBase, vbTextCompare) = 0 Then
        ToLogPath = "."
        Exit Function
    End If

    baseWithSep = normalizedBase
    If Right$(baseWithSep, 1) <> sep Then baseWithSep = baseWithSep & sep

    If Len(normalizedPath) > Len(baseWithSep) Then
        If StrComp(Left$(normalizedPath, Len(baseWithSep)), baseWithSep, vbTextCompare) = 0 Then
            ToLogPath = Mid$(normalizedPath, Len(baseWithSep) + 1)
            Exit Function
        End If
    End If

    ToLogPath = pathValue
End Function

Private Function SanitizeLogText(ByVal textValue As String) As String
    Dim result As String
    Dim basePath As String
    Dim baseWithSep As String
    Dim basePathSlash As String
    Dim baseWithSepSlash As String
    Dim sep As String

    result = textValue
    basePath = ThisWorkbook.Path

    If Len(basePath) > 0 Then
        sep = Application.PathSeparator

        baseWithSep = basePath
        If Right$(baseWithSep, 1) <> sep Then baseWithSep = baseWithSep & sep

        basePathSlash = Replace(basePath, "\", "/")
        baseWithSepSlash = Replace(baseWithSep, "\", "/")

        result = Replace(result, baseWithSep, "", 1, -1, vbTextCompare)
        result = Replace(result, basePath, ".", 1, -1, vbTextCompare)
        result = Replace(result, baseWithSepSlash, "", 1, -1, vbTextCompare)
        result = Replace(result, basePathSlash, ".", 1, -1, vbTextCompare)
    End If

    SanitizeLogText = result
End Function
```
