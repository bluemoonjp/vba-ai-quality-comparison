Option Explicit

Private Const SHEET_MONTHLY As String = "月次集計"
Private Const SHEET_DAILY As String = "日次集計"
Private Const SHEET_MASTER As String = "マスタ"
Private Const SHEET_LOG As String = "異常系"

Private Const COL_ROW_TYPE As Long = 1
Private Const COL_GROUP As Long = 2
Private Const COL_BRANCH_CODE As Long = 4
Private Const COL_BRANCH_NAME As Long = 5
Private Const COL_BIZ_CODE As Long = 6
Private Const COL_BIZ_NAME As Long = 7
Private Const COL_DAY_START As Long = 8      ' H
Private Const COL_DAY_END As Long = 38       ' AL
Private Const COL_MONTH_TOTAL As Long = 39   ' AM

Private Const DAILY_COL_STATUS As Long = 1
Private Const DAILY_COL_BRANCH_CODE As Long = 2
Private Const DAILY_COL_BRANCH_NAME As Long = 3
Private Const DAILY_COL_BIZ_CODE As Long = 4
Private Const DAILY_COL_BIZ_NAME As Long = 5
Private Const DAILY_COL_COUNT As Long = 6

Public Sub RunMonthlyClaimAggregation()
    On Error GoTo FatalError

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim yymm As String
    yymm = GetMonthlyYYMM(wb.Name)

    If yymm = "" Then
        MsgBox "月次ブック名が想定形式ではありません: " & wb.Name, vbCritical
        Exit Sub
    End If

    Dim basePath As String
    basePath = wb.Path & Application.PathSeparator & "daily" & Application.PathSeparator & yymm

    Dim wsMonthly As Worksheet
    Dim wsLog As Worksheet
    Set wsMonthly = wb.Worksheets(SHEET_MONTHLY)
    Set wsLog = wb.Worksheets(SHEET_LOG)

    InitLog wsLog
    LogIssue wsLog, "INFO", "START", "", 0, "", "", "", "", "", "月次集計を開始: YYMM=" & yymm

    Dim lastDay As Long
    lastDay = Day(DateSerial(2000 + CLng(Left$(yymm, 2)), CLng(Right$(yymm, 2)) + 1, 0))

    Dim detailRowByKey As Object
    Dim subtotalRows As Collection
    Set detailRowByKey = CreateObject("Scripting.Dictionary")
    Set subtotalRows = New Collection

    BuildMonthlyIndex wsMonthly, detailRowByKey, subtotalRows, wsLog

    Dim knownBranches As Object
    Dim knownBiz As Object
    Dim validPairs As Object
    Set knownBranches = CreateObject("Scripting.Dictionary")
    Set knownBiz = CreateObject("Scripting.Dictionary")
    Set validPairs = CreateObject("Scripting.Dictionary")

    BuildMasterCodes wb, wsMonthly, knownBranches, knownBiz, validPairs, wsLog

    ClearDetailDayValues wsMonthly

    If Dir(basePath, vbDirectory) = "" Then
        LogIssue wsLog, "ERROR", "FOLDER_NOT_FOUND", RelativePath(basePath), 0, "", "", "", "", "", "日次フォルダが存在しません。"
        GoTo Finish
    End If

    Dim dateFiles As Object
    Set dateFiles = CreateObject("Scripting.Dictionary")

    CollectDailyFiles basePath, yymm, dateFiles, wsLog

    Dim dateKey As Variant
    For Each dateKey In dateFiles.Keys
        ProcessOneDailyFilePerDate CStr(dateKey), dateFiles(dateKey), wsMonthly, wsLog, _
                                   detailRowByKey, knownBranches, knownBiz, validPairs
    Next dateKey

    RecalculateDetailTotals wsMonthly, lastDay
    RecalculateSubtotals wsMonthly, subtotalRows, lastDay

Finish:
    LogIssue wsLog, "INFO", "END", "", 0, "", "", "", "", "", "月次集計を終了しました。"
    MsgBox "月次クレーム集計が完了しました。異常系シートを確認してください。", vbInformation
    Exit Sub

FatalError:
    On Error Resume Next
    LogIssue ThisWorkbook.Worksheets(SHEET_LOG), "FATAL", "UNEXPECTED_ERROR", "", 0, "", "", "", "", "", Err.Number & ": " & Err.Description
    MsgBox "処理中に予期しないエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Sub ProcessOneDailyFilePerDate(ByVal yymmdd As String, ByVal files As Collection, _
                                       ByVal wsMonthly As Worksheet, ByVal wsLog As Worksheet, _
                                       ByVal detailRowByKey As Object, ByVal knownBranches As Object, _
                                       ByVal knownBiz As Object, ByVal validPairs As Object)
    Dim selectedPath As String
    selectedPath = GetMinPath(files)

    Dim i As Long
    For i = 1 To files.Count
        If CStr(files(i)) <> selectedPath Then
            LogIssue wsLog, "WARN", "DUPLICATE_DAILY_FILE", RelativePath(CStr(files(i))), 0, yymmdd, "", "", "", "", _
                     "同じ日付の日次ファイルが複数あります。フルパス昇順の先頭1件のみ処理し、このファイルはスキップします。処理対象=" & RelativePath(selectedPath)
        End If
    Next i

    ProcessDailyFile selectedPath, yymmdd, wsMonthly, wsLog, detailRowByKey, knownBranches, knownBiz, validPairs
End Sub

Private Function GetMinPath(ByVal files As Collection) As String
    Dim i As Long
    Dim minPath As String

    minPath = CStr(files(1))
    For i = 2 To files.Count
        If StrComp(CStr(files(i)), minPath, vbTextCompare) < 0 Then
            minPath = CStr(files(i))
        End If
    Next i

    GetMinPath = minPath
End Function

Private Function GetMonthlyYYMM(ByVal fileName As String) As String
    If fileName Like "月次クレーム集計####.xlsm" Then
        GetMonthlyYYMM = Mid$(fileName, Len("月次クレーム集計") + 1, 4)
        If Not IsValidYYMM(GetMonthlyYYMM) Then GetMonthlyYYMM = ""
    End If
End Function

Private Function IsValidYYMM(ByVal yymm As String) As Boolean
    On Error GoTo Invalid

    If Len(yymm) <> 4 Or Not IsNumeric(yymm) Then Exit Function

    Dim mm As Long
    mm = CLng(Right$(yymm, 2))

    IsValidYYMM = (mm >= 1 And mm <= 12)
    Exit Function

Invalid:
    IsValidYYMM = False
End Function

Private Function GetDailyDateKey(ByVal fileName As String) As String
    If fileName Like "クレーム集計######.xlsx" Then
        GetDailyDateKey = Mid$(fileName, Len("クレーム集計") + 1, 6)
        If Not IsNumeric(GetDailyDateKey) Then GetDailyDateKey = ""
    End If
End Function

Private Function IsValidYYMMDD(ByVal yymmdd As String) As Boolean
    On Error GoTo Invalid

    If Len(yymmdd) <> 6 Or Not IsNumeric(yymmdd) Then Exit Function

    Dim yy As Long
    Dim mm As Long
    Dim dd As Long
    yy = CLng(Left$(yymmdd, 2))
    mm = CLng(Mid$(yymmdd, 3, 2))
    dd = CLng(Right$(yymmdd, 2))

    Dim d As Date
    d = DateSerial(2000 + yy, mm, dd)

    IsValidYYMMDD = (Year(d) = 2000 + yy And Month(d) = mm And Day(d) = dd)
    Exit Function

Invalid:
    IsValidYYMMDD = False
End Function

Private Sub CollectDailyFiles(ByVal folderPath As String, ByVal targetYYMM As String, _
                              ByVal dateFiles As Object, ByVal wsLog As Worksheet)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "xlsx" Then
            Dim yymmdd As String
            yymmdd = GetDailyDateKey(file.Name)

            If yymmdd = "" Then
                LogIssue wsLog, "WARN", "INVALID_DAILY_FILENAME", RelativePath(file.Path), 0, "", "", "", "", "", "日次ファイル名形式が違います。"
            ElseIf Not IsValidYYMMDD(yymmdd) Then
                LogIssue wsLog, "WARN", "INVALID_DAILY_DATE", RelativePath(file.Path), 0, yymmdd, "", "", "", "", "日次ファイル名の日付が不正です。"
            ElseIf Left$(yymmdd, 4) <> targetYYMM Then
                LogIssue wsLog, "WARN", "OUT_OF_MONTH_FILE", RelativePath(file.Path), 0, yymmdd, "", "", "", "", "対象月と異なる日次ファイルです。"
            Else
                If Not dateFiles.Exists(yymmdd) Then
                    dateFiles.Add yymmdd, New Collection
                End If
                dateFiles(yymmdd).Add CStr(file.Path)
            End If
        End If
    Next file

    For Each subFolder In folder.SubFolders
        CollectDailyFiles subFolder.Path, targetYYMM, dateFiles, wsLog
    Next subFolder
End Sub

Private Sub ProcessDailyFile(ByVal filePath As String, ByVal yymmdd As String, _
                             ByVal wsMonthly As Worksheet, ByVal wsLog As Worksheet, _
                             ByVal detailRowByKey As Object, ByVal knownBranches As Object, _
                             ByVal knownBiz As Object, ByVal validPairs As Object)
    On Error GoTo FileError

    Dim wbDaily As Workbook
    Set wbDaily = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=False)

    Dim wsDaily As Worksheet
    On Error Resume Next
    Set wsDaily = wbDaily.Worksheets(SHEET_DAILY)
    On Error GoTo FileError

    If wsDaily Is Nothing Then
        LogIssue wsLog, "ERROR", "DAILY_SHEET_NOT_FOUND", RelativePath(filePath), 0, yymmdd, "", "", "", "", "日次集計シートがありません。"
        wbDaily.Close SaveChanges:=False
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsDaily.Cells(wsDaily.Rows.Count, DAILY_COL_STATUS).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        ProcessDailyRow wsDaily, r, yymmdd, filePath, wsMonthly, wsLog, detailRowByKey, knownBranches, knownBiz, validPairs
    Next r

    wbDaily.Close SaveChanges:=False
    Exit Sub

FileError:
    LogIssue wsLog, "ERROR", "DAILY_FILE_OPEN_OR_READ_ERROR", RelativePath(filePath), 0, yymmdd, "", "", "", "", Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wbDaily Is Nothing Then wbDaily.Close SaveChanges:=False
End Sub

Private Sub ProcessDailyRow(ByVal wsDaily As Worksheet, ByVal rowNumber As Long, _
                            ByVal yymmdd As String, ByVal filePath As String, _
                            ByVal wsMonthly As Worksheet, ByVal wsLog As Worksheet, _
                            ByVal detailRowByKey As Object, ByVal knownBranches As Object, _
                            ByVal knownBiz As Object, ByVal validPairs As Object)
    Dim rawStatus As String
    Dim normalizedStatus As String
    Dim branchCode As String
    Dim bizCode As String
    Dim claimValue As Variant

    rawStatus = Trim$(CStr(wsDaily.Cells(rowNumber, DAILY_COL_STATUS).Value))
    normalizedStatus = LCase$(rawStatus)
    branchCode = Trim$(CStr(wsDaily.Cells(rowNumber, DAILY_COL_BRANCH_CODE).Value))
    bizCode = Trim$(CStr(wsDaily.Cells(rowNumber, DAILY_COL_BIZ_CODE).Value))
    claimValue = wsDaily.Cells(rowNumber, DAILY_COL_COUNT).Value

    Dim isValid As Boolean
    isValid = ValidateDailyRow(filePath, rowNumber, yymmdd, rawStatus, branchCode, bizCode, claimValue, _
                               wsLog, detailRowByKey, knownBranches, knownBiz, validPairs)

    If normalizedStatus <> "ok" Then
        LogIssue wsLog, "WARN", "NON_OK_ROW", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, claimValue, _
                 "処理区分が ok ではないため転記しません。"
        Exit Sub
    End If

    If Not isValid Then Exit Sub

    Dim key As String
    Dim targetRow As Long
    Dim targetCol As Long
    Dim dayNum As Long

    key = MakeKey(branchCode, bizCode)
    targetRow = CLng(detailRowByKey(key))
    dayNum = CLng(Right$(yymmdd, 2))
    targetCol = COL_DAY_START + dayNum - 1

    wsMonthly.Cells(targetRow, targetCol).Value = NzNumber(wsMonthly.Cells(targetRow, targetCol).Value) + CDbl(claimValue)
End Sub

Private Function ValidateDailyRow(ByVal filePath As String, ByVal rowNumber As Long, _
                                  ByVal yymmdd As String, ByVal rawStatus As String, _
                                  ByVal branchCode As String, ByVal bizCode As String, _
                                  ByVal claimValue As Variant, ByVal wsLog As Worksheet, _
                                  ByVal detailRowByKey As Object, ByVal knownBranches As Object, _
                                  ByVal knownBiz As Object, ByVal validPairs As Object) As Boolean
    Dim hasError As Boolean
    Dim key As String

    key = MakeKey(branchCode, bizCode)

    If branchCode = "" Or Not knownBranches.Exists(branchCode) Then
        LogIssue wsLog, "ERROR", "UNKNOWN_BRANCH_CODE", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, claimValue, "未知の支店コードです。"
        hasError = True
    End If

    If bizCode = "" Or Not knownBiz.Exists(bizCode) Then
        LogIssue wsLog, "ERROR", "UNKNOWN_BIZ_CODE", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, claimValue, "未知の業務コードです。"
        hasError = True
    End If

    If Trim$(CStr(claimValue)) = "" Then
        LogIssue wsLog, "ERROR", "BLANK_CLAIM_COUNT", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, "", "クレーム件数が空欄です。"
        hasError = True
    ElseIf Not IsNumeric(claimValue) Then
        LogIssue wsLog, "ERROR", "NON_NUMERIC_CLAIM_COUNT", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, claimValue, "クレーム件数が数値ではありません。"
        hasError = True
    End If

    If Not hasError Then
        If validPairs.Count > 0 Then
            If Not validPairs.Exists(key) Then
                LogIssue wsLog, "ERROR", "INVALID_BRANCH_BIZ_PAIR", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, claimValue, "マスタ上、有効な支店コード + 業務コードの組み合わせではありません。"
                hasError = True
            End If
        End If
    End If

    If Not hasError Then
        If Not detailRowByKey.Exists(key) Then
            LogIssue wsLog, "ERROR", "MONTHLY_KEY_NOT_FOUND", RelativePath(filePath), rowNumber, yymmdd, branchCode, bizCode, rawStatus, claimValue, "月次表に該当する DETAIL 行がありません。"
            hasError = True
        End If
    End If

    ValidateDailyRow = Not hasError
End Function

Private Sub BuildMonthlyIndex(ByVal ws As Worksheet, ByVal detailRowByKey As Object, _
                              ByVal subtotalRows As Collection, ByVal wsLog As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim rowType As String
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))

        If rowType = "DETAIL" Then
            Dim branchCode As String
            Dim bizCode As String
            Dim key As String

            branchCode = Trim$(CStr(ws.Cells(r, COL_BRANCH_CODE).Value))
            bizCode = Trim$(CStr(ws.Cells(r, COL_BIZ_CODE).Value))
            key = MakeKey(branchCode, bizCode)

            If branchCode <> "" And bizCode <> "" Then
                If detailRowByKey.Exists(key) Then
                    LogIssue wsLog, "ERROR", "DUPLICATE_MONTHLY_DETAIL_KEY", "", r, "", branchCode, bizCode, "", "", "月次集計に同じ DETAIL キーが複数あります。先勝ちで処理します。"
                Else
                    detailRowByKey.Add key, r
                End If
            End If

        ElseIf rowType = "SUBTOTAL" Then
            subtotalRows.Add r
        End If
    Next r
End Sub

Private Sub BuildMasterCodes(ByVal wb As Workbook, ByVal wsMonthly As Worksheet, _
                             ByVal knownBranches As Object, ByVal knownBiz As Object, _
                             ByVal validPairs As Object, ByVal wsLog As Worksheet)
    Dim wsMaster As Worksheet
    On Error Resume Next
    Set wsMaster = wb.Worksheets(SHEET_MASTER)
    On Error GoTo 0

    If Not wsMaster Is Nothing Then
        ReadMasterSheet wsMaster, knownBranches, knownBiz, validPairs, wsLog
    Else
        LogIssue wsLog, "WARN", "MASTER_SHEET_NOT_FOUND", "", 0, "", "", "", "", "", "マスタシートがありません。月次 DETAIL 行から既知コードと有効組み合わせを補完します。"
    End If

    AddMasterCodesFromMonthlyDetail wsMonthly, knownBranches, knownBiz, validPairs

    If knownBranches.Count = 0 Then
        LogIssue wsLog, "WARN", "MASTER_BRANCH_EMPTY", "", 0, "", "", "", "", "", "支店コードを取得できませんでした。"
    End If

    If knownBiz.Count = 0 Then
        LogIssue wsLog, "WARN", "MASTER_BIZ_EMPTY", "", 0, "", "", "", "", "", "業務コードを取得できませんでした。"
    End If

    If validPairs.Count = 0 Then
        LogIssue wsLog, "WARN", "MASTER_PAIR_EMPTY", "", 0, "", "", "", "", "", "支店+業務の有効組み合わせを取得できませんでした。"
    End If
End Sub

Private Sub ReadMasterSheet(ByVal ws As Worksheet, ByVal knownBranches As Object, _
                            ByVal knownBiz As Object, ByVal validPairs As Object, _
                            ByVal wsLog As Worksheet)
    Dim branchCol As Long
    Dim bizCol As Long

    branchCol = FindHeaderColumnAtRow(ws, 2, "支店コード")
    If branchCol > 0 Then
        AddCodesBetweenRows ws, branchCol, 3, 11, knownBranches
    Else
        LogIssue wsLog, "WARN", "MASTER_BRANCH_HEADER_NOT_FOUND", "", 0, "", "", "", "", "", "マスタ2行目に支店コード見出しが見つかりません。"
    End If

    bizCol = FindHeaderColumnAtRow(ws, 12, "業務コード")
    If bizCol > 0 Then
        AddCodesBetweenRows ws, bizCol, 13, 23, knownBiz
    Else
        LogIssue wsLog, "WARN", "MASTER_BIZ_HEADER_NOT_FOUND", "", 0, "", "", "", "", "", "マスタ12行目に業務コード見出しが見つかりません。"
    End If

    Dim pairBranchCol As Long
    Dim pairBizCol As Long
    pairBranchCol = FindHeaderColumnAtRow(ws, 24, "支店コード")
    pairBizCol = FindHeaderColumnAtRow(ws, 24, "業務コード")

    If pairBranchCol > 0 And pairBizCol > 0 Then
        AddPairsFromRows ws, pairBranchCol, pairBizCol, 25, ws.Cells(ws.Rows.Count, pairBranchCol).End(xlUp).Row, _
                         knownBranches, knownBiz, validPairs
    Else
        LogIssue wsLog, "WARN", "MASTER_PAIR_HEADER_NOT_FOUND", "", 0, "", "", "", "", "", "マスタ24行目に支店コード/業務コード見出しが見つかりません。"
    End If
End Sub

Private Function FindHeaderColumnAtRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(headerRow, c).Value)) = headerText Then
            FindHeaderColumnAtRow = c
            Exit Function
        End If
    Next c
End Function

Private Sub AddCodesBetweenRows(ByVal ws As Worksheet, ByVal colNum As Long, _
                                ByVal startRow As Long, ByVal endRow As Long, ByVal dict As Object)
    Dim r As Long
    Dim code As String

    For r = startRow To endRow
        code = Trim$(CStr(ws.Cells(r, colNum).Value))
        If code <> "" Then
            If Not dict.Exists(code) Then dict.Add code, True
        End If
    Next r
End Sub

Private Sub AddPairsFromRows(ByVal ws As Worksheet, ByVal branchCol As Long, ByVal bizCol As Long, _
                             ByVal startRow As Long, ByVal endRow As Long, _
                             ByVal knownBranches As Object, ByVal knownBiz As Object, ByVal validPairs As Object)
    Dim r As Long
    For r = startRow To endRow
        Dim branchCode As String
        Dim bizCode As String
        Dim key As String

        branchCode = Trim$(CStr(ws.Cells(r, branchCol).Value))
        bizCode = Trim$(CStr(ws.Cells(r, bizCol).Value))

        If branchCode <> "" And bizCode <> "" Then
            key = MakeKey(branchCode, bizCode)

            If Not knownBranches.Exists(branchCode) Then knownBranches.Add branchCode, True
            If Not knownBiz.Exists(bizCode) Then knownBiz.Add bizCode, True
            If Not validPairs.Exists(key) Then validPairs.Add key, True
        End If
    Next r
End Sub

Private Sub AddMasterCodesFromMonthlyDetail(ByVal ws As Worksheet, ByVal knownBranches As Object, _
                                            ByVal knownBiz As Object, ByVal validPairs As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value))) = "DETAIL" Then
            Dim branchCode As String
            Dim bizCode As String
            Dim key As String

            branchCode = Trim$(CStr(ws.Cells(r, COL_BRANCH_CODE).Value))
            bizCode = Trim$(CStr(ws.Cells(r, COL_BIZ_CODE).Value))
            key = MakeKey(branchCode, bizCode)

            If branchCode <> "" Then
                If Not knownBranches.Exists(branchCode) Then knownBranches.Add branchCode, True
            End If

            If bizCode <> "" Then
                If Not knownBiz.Exists(bizCode) Then knownBiz.Add bizCode, True
            End If

            If branchCode <> "" And bizCode <> "" Then
                If Not validPairs.Exists(key) Then validPairs.Add key, True
            End If
        End If
    Next r
End Sub

Private Sub ClearDetailDayValues(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value))) = "DETAIL" Then
            ws.Range(ws.Cells(r, COL_DAY_START), ws.Cells(r, COL_DAY_END)).ClearContents
        End If
    Next r
End Sub

Private Sub RecalculateDetailTotals(ByVal ws As Worksheet, ByVal lastDay As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    Dim c As Long
    Dim total As Double

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value))) = "DETAIL" Then
            ws.Cells(r, COL_MONTH_TOTAL).ClearContents

            total = 0
            For c = COL_DAY_START To COL_DAY_START + lastDay - 1
                total = total + NzNumber(ws.Cells(r, c).Value)
            Next c

            If total <> 0 Then ws.Cells(r, COL_MONTH_TOTAL).Value = total

            If lastDay < 31 Then
                ws.Range(ws.Cells(r, COL_DAY_START + lastDay), ws.Cells(r, COL_DAY_END)).ClearContents
            End If
        End If
    Next r
End Sub

Private Sub RecalculateSubtotals(ByVal ws As Worksheet, ByVal subtotalRows As Collection, ByVal lastDay As Long)
    Dim subtotalRow As Variant

    For Each subtotalRow In subtotalRows
        Dim groupName As String
        groupName = Trim$(CStr(ws.Cells(CLng(subtotalRow), COL_GROUP).Value))

        ws.Range(ws.Cells(CLng(subtotalRow), COL_DAY_START), ws.Cells(CLng(subtotalRow), COL_DAY_END)).ClearContents
        ws.Cells(CLng(subtotalRow), COL_MONTH_TOTAL).ClearContents

        Dim c As Long
        For c = COL_DAY_START To COL_DAY_START + lastDay - 1
            Dim dayTotal As Double
            dayTotal = SumDetailByGroup(ws, groupName, c)
            If dayTotal <> 0 Then ws.Cells(CLng(subtotalRow), c).Value = dayTotal
        Next c

        Dim monthTotal As Double
        monthTotal = SumDetailByGroup(ws, groupName, COL_MONTH_TOTAL)
        If monthTotal <> 0 Then ws.Cells(CLng(subtotalRow), COL_MONTH_TOTAL).Value = monthTotal

        If lastDay < 31 Then
            ws.Range(ws.Cells(CLng(subtotalRow), COL_DAY_START + lastDay), ws.Cells(CLng(subtotalRow), COL_DAY_END)).ClearContents
        End If
    Next subtotalRow
End Sub

Private Function SumDetailByGroup(ByVal ws As Worksheet, ByVal groupName As String, ByVal valueCol As Long) As Double
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    Dim total As Double

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value))) = "DETAIL" Then
            If Trim$(CStr(ws.Cells(r, COL_GROUP).Value)) = groupName Then
                total = total + NzNumber(ws.Cells(r, valueCol).Value)
            End If
        End If
    Next r

    SumDetailByGroup = total
End Function

Private Function MakeKey(ByVal branchCode As String, ByVal bizCode As String) As String
    MakeKey = Trim$(branchCode) & "|" & Trim$(bizCode)
End Function

Private Function NzNumber(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        NzNumber = CDbl(v)
    Else
        NzNumber = 0
    End If
End Function

Private Function RelativePath(ByVal fullPath As String) As String
    Dim basePath As String
    basePath = ThisWorkbook.Path

    If fullPath = "" Then
        RelativePath = ""
    ElseIf StrComp(Left$(fullPath, Len(basePath)), basePath, vbTextCompare) = 0 Then
        RelativePath = Mid$(fullPath, Len(basePath) + 2)
    Else
        RelativePath = fullPath
    End If
End Function

Private Sub InitLog(ByVal ws As Worksheet)
    ws.Cells.ClearContents

    ws.Range("A1:K1").Value = Array( _
        "記録日時", "重要度", "分類", "ファイル", "行番号", _
        "日付", "支店コード", "業務コード", "処理区分", "値", "内容" _
    )
End Sub

Private Sub LogIssue(ByVal ws As Worksheet, ByVal severity As String, ByVal category As String, _
                     ByVal filePath As String, ByVal rowNumber As Long, ByVal dateKey As String, _
                     ByVal branchCode As String, ByVal bizCode As String, ByVal processStatus As String, _
                     ByVal rawValue As Variant, ByVal message As String)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = severity
    ws.Cells(r, 3).Value = category
    ws.Cells(r, 4).Value = filePath
    If rowNumber > 0 Then ws.Cells(r, 5).Value = rowNumber
    ws.Cells(r, 6).Value = dateKey
    ws.Cells(r, 7).Value = branchCode
    ws.Cells(r, 8).Value = bizCode
    ws.Cells(r, 9).Value = processStatus
    ws.Cells(r, 10).Value = rawValue
    ws.Cells(r, 11).Value = message
End Sub
