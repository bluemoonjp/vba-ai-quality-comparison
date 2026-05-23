# chatgpt-instant 初回出力

```vba
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

    Dim wsMonthly As Worksheet, wsLog As Worksheet
    Set wsMonthly = wb.Worksheets(SHEET_MONTHLY)
    Set wsLog = wb.Worksheets(SHEET_LOG)

    InitLog wsLog
    LogIssue wsLog, "INFO", "START", "", 0, "", "", "", "", "月次集計を開始: YYMM=" & yymm

    Dim lastDay As Long
    lastDay = Day(DateSerial(2000 + CInt(Left$(yymm, 2)), CInt(Right$(yymm, 2)) + 1, 0))

    Dim detailRowByKey As Object
    Dim subtotalRows As Collection
    Set detailRowByKey = CreateObject("Scripting.Dictionary")
    Set subtotalRows = New Collection

    BuildMonthlyIndex wsMonthly, detailRowByKey, subtotalRows, wsLog
    ClearMonthlyValues wsMonthly, lastDay

    Dim knownBranches As Object, knownBiz As Object
    Set knownBranches = CreateObject("Scripting.Dictionary")
    Set knownBiz = CreateObject("Scripting.Dictionary")
    BuildKnownCodes wb, wsMonthly, knownBranches, knownBiz, wsLog

    If Dir(basePath, vbDirectory) = "" Then
        LogIssue wsLog, "ERROR", "FOLDER_NOT_FOUND", basePath, 0, "", "", "", "", "日次フォルダが存在しません。"
        GoTo Finish
    End If

    Dim dateFiles As Object
    Set dateFiles = CreateObject("Scripting.Dictionary")

    CollectDailyFiles basePath, yymm, dateFiles, wsLog

    Dim dateKey As Variant
    For Each dateKey In dateFiles.Keys
        If dateFiles(dateKey).Count > 1 Then
            Dim p As Variant
            For Each p In dateFiles(dateKey)
                LogIssue wsLog, "ERROR", "DUPLICATE_DAILY_FILE", CStr(p), 0, CStr(dateKey), "", "", "", "同じ日付の日次ファイルが複数あります。この日付は転記しません。"
            Next p
        Else
            ProcessDailyFile CStr(dateFiles(dateKey)(1)), CStr(dateKey), wsMonthly, wsLog, _
                             detailRowByKey, knownBranches, knownBiz
        End If
    Next dateKey

    RecalculateDetailTotals wsMonthly, lastDay
    RecalculateSubtotals wsMonthly, subtotalRows, lastDay

Finish:
    LogIssue wsLog, "INFO", "END", "", 0, "", "", "", "", "月次集計を終了しました。"
    MsgBox "月次クレーム集計が完了しました。異常系シートを確認してください。", vbInformation
    Exit Sub

FatalError:
    On Error Resume Next
    LogIssue ThisWorkbook.Worksheets(SHEET_LOG), "FATAL", "UNEXPECTED_ERROR", "", 0, "", "", "", "", Err.Number & ": " & Err.Description
    MsgBox "処理中に予期しないエラーが発生しました: " & Err.Description, vbCritical
End Sub

Private Function GetMonthlyYYMM(ByVal fileName As String) As String
    If fileName Like "月次クレーム集計####.xlsm" Then
        GetMonthlyYYMM = Mid$(fileName, Len("月次クレーム集計") + 1, 4)
        If Not IsValidYYMM(GetMonthlyYYMM) Then GetMonthlyYYMM = ""
    End If
End Function

Private Function IsValidYYMM(ByVal yymm As String) As Boolean
    On Error GoTo Invalid
    If Len(yymm) <> 4 Or Not IsNumeric(yymm) Then Exit Function
    Dim yy As Long, mm As Long
    yy = CLng(Left$(yymm, 2))
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

    Dim yy As Long, mm As Long, dd As Long
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
    Dim fso As Object, folder As Object, subFolder As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "xlsx" Then
            Dim yymmdd As String
            yymmdd = GetDailyDateKey(file.Name)

            If yymmdd = "" Then
                LogIssue wsLog, "WARN", "INVALID_DAILY_FILENAME", file.Path, 0, "", "", "", "", "日次ファイル名形式が違います。"
            ElseIf Not IsValidYYMMDD(yymmdd) Then
                LogIssue wsLog, "WARN", "INVALID_DAILY_DATE", file.Path, 0, yymmdd, "", "", "", "日次ファイル名の日付が不正です。"
            ElseIf Left$(yymmdd, 4) <> targetYYMM Then
                LogIssue wsLog, "WARN", "OUT_OF_MONTH_FILE", file.Path, 0, yymmdd, "", "", "", "対象月と異なる日次ファイルです。"
            Else
                If Not dateFiles.Exists(yymmdd) Then
                    dateFiles.Add yymmdd, New Collection
                End If
                dateFiles(yymmdd).Add file.Path
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
                             ByVal knownBiz As Object)
    On Error GoTo FileError

    Dim wbDaily As Workbook
    Set wbDaily = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=False)

    Dim wsDaily As Worksheet
    On Error Resume Next
    Set wsDaily = wbDaily.Worksheets(SHEET_DAILY)
    On Error GoTo FileError

    If wsDaily Is Nothing Then
        LogIssue wsLog, "ERROR", "DAILY_SHEET_NOT_FOUND", filePath, 0, yymmdd, "", "", "", "日次集計シートがありません。"
        wbDaily.Close SaveChanges:=False
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsDaily.Cells(wsDaily.Rows.Count, DAILY_COL_STATUS).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim status As String
        status = LCase$(Trim$(CStr(wsDaily.Cells(r, DAILY_COL_STATUS).Value)))

        Dim branchCode As String, bizCode As String
        branchCode = Trim$(CStr(wsDaily.Cells(r, DAILY_COL_BRANCH_CODE).Value))
        bizCode = Trim$(CStr(wsDaily.Cells(r, DAILY_COL_BIZ_CODE).Value))

        If status <> "ok" Then
            LogIssue wsLog, "WARN", "NON_OK_ROW", filePath, r, yymmdd, branchCode, bizCode, "", "処理区分が ok ではないため転記しません。"
        Else
            Dim claimValue As Variant
            claimValue = wsDaily.Cells(r, DAILY_COL_COUNT).Value

            If branchCode = "" Or Not knownBranches.Exists(branchCode) Then
                LogIssue wsLog, "ERROR", "UNKNOWN_BRANCH_CODE", filePath, r, yymmdd, branchCode, bizCode, claimValue, "未知の支店コードです。"
            ElseIf bizCode = "" Or Not knownBiz.Exists(bizCode) Then
                LogIssue wsLog, "ERROR", "UNKNOWN_BIZ_CODE", filePath, r, yymmdd, branchCode, bizCode, claimValue, "未知の業務コードです。"
            ElseIf Trim$(CStr(claimValue)) = "" Then
                LogIssue wsLog, "ERROR", "BLANK_CLAIM_COUNT", filePath, r, yymmdd, branchCode, bizCode, "", "クレーム件数が空欄です。"
            ElseIf Not IsNumeric(claimValue) Then
                LogIssue wsLog, "ERROR", "NON_NUMERIC_CLAIM_COUNT", filePath, r, yymmdd, branchCode, bizCode, claimValue, "クレーム件数が数値ではありません。"
            Else
                Dim key As String
                key = MakeKey(branchCode, bizCode)

                If Not detailRowByKey.Exists(key) Then
                    LogIssue wsLog, "ERROR", "MONTHLY_KEY_NOT_FOUND", filePath, r, yymmdd, branchCode, bizCode, claimValue, "月次表に該当する DETAIL 行がありません。"
                Else
                    Dim dayNum As Long, targetCol As Long, targetRow As Long
                    dayNum = CLng(Right$(yymmdd, 2))
                    targetCol = COL_DAY_START + dayNum - 1
                    targetRow = CLng(detailRowByKey(key))

                    wsMonthly.Cells(targetRow, targetCol).Value = NzNumber(wsMonthly.Cells(targetRow, targetCol).Value) + CDbl(claimValue)
                End If
            End If
        End If
    Next r

    wbDaily.Close SaveChanges:=False
    Exit Sub

FileError:
    LogIssue wsLog, "ERROR", "DAILY_FILE_OPEN_OR_READ_ERROR", filePath, 0, yymmdd, "", "", "", Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wbDaily Is Nothing Then wbDaily.Close SaveChanges:=False
End Sub

Private Sub BuildMonthlyIndex(ByVal ws As Worksheet, ByVal detailRowByKey As Object, _
                              ByVal subtotalRows As Collection, ByVal wsLog As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim rowType As String
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))

        If rowType = "DETAIL" Then
            Dim branchCode As String, bizCode As String, key As String
            branchCode = Trim$(CStr(ws.Cells(r, COL_BRANCH_CODE).Value))
            bizCode = Trim$(CStr(ws.Cells(r, COL_BIZ_CODE).Value))
            key = MakeKey(branchCode, bizCode)

            If branchCode <> "" And bizCode <> "" Then
                If detailRowByKey.Exists(key) Then
                    LogIssue wsLog, "ERROR", "DUPLICATE_MONTHLY_DETAIL_KEY", "", r, "", branchCode, bizCode, "", "月次集計に同じ DETAIL キーが複数あります。先勝ちで処理します。"
                Else
                    detailRowByKey.Add key, r
                End If
            End If
        ElseIf rowType = "SUBTOTAL" Then
            subtotalRows.Add r
        End If
    Next r
End Sub

Private Sub BuildKnownCodes(ByVal wb As Workbook, ByVal wsMonthly As Worksheet, _
                            ByVal knownBranches As Object, ByVal knownBiz As Object, _
                            ByVal wsLog As Worksheet)
    Dim wsMaster As Worksheet
    On Error Resume Next
    Set wsMaster = wb.Worksheets(SHEET_MASTER)
    On Error GoTo 0

    If Not wsMaster Is Nothing Then
        Dim branchCol As Long, bizCol As Long
        branchCol = FindHeaderColumn(wsMaster, "支店コード")
        bizCol = FindHeaderColumn(wsMaster, "業務コード")

        If branchCol > 0 Then AddCodesFromColumn wsMaster, branchCol, knownBranches
        If bizCol > 0 Then AddCodesFromColumn wsMaster, bizCol, knownBiz
    End If

    AddKnownCodesFromMonthly wsMonthly, knownBranches, knownBiz

    If knownBranches.Count = 0 Then LogIssue wsLog, "WARN", "MASTER_BRANCH_EMPTY", "", 0, "", "", "", "", "支店コードマスタを取得できませんでした。"
    If knownBiz.Count = 0 Then LogIssue wsLog, "WARN", "MASTER_BIZ_EMPTY", "", 0, "", "", "", "", "業務コードマスタを取得できませんでした。"
End Sub

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerText Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Sub AddCodesFromColumn(ByVal ws As Worksheet, ByVal colNum As Long, ByVal dict As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row

    Dim r As Long, code As String
    For r = 2 To lastRow
        code = Trim$(CStr(ws.Cells(r, colNum).Value))
        If code <> "" Then
            If Not dict.Exists(code) Then dict.Add code, True
        End If
    Next r
End Sub

Private Sub AddKnownCodesFromMonthly(ByVal ws As Worksheet, ByVal knownBranches As Object, ByVal knownBiz As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value))) = "DETAIL" Then
            Dim branchCode As String, bizCode As String
            branchCode = Trim$(CStr(ws.Cells(r, COL_BRANCH_CODE).Value))
            bizCode = Trim$(CStr(ws.Cells(r, COL_BIZ_CODE).Value))

            If branchCode <> "" Then If Not knownBranches.Exists(branchCode) Then knownBranches.Add branchCode, True
            If bizCode <> "" Then If Not knownBiz.Exists(bizCode) Then knownBiz.Add bizCode, True
        End If
    Next r
End Sub

Private Sub ClearMonthlyValues(ByVal ws As Worksheet, ByVal lastDay As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim rowType As String
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))

        If rowType = "DETAIL" Or rowType = "SUBTOTAL" Then
            ws.Range(ws.Cells(r, COL_DAY_START), ws.Cells(r, COL_DAY_END)).ClearContents
            ws.Cells(r, COL_MONTH_TOTAL).ClearContents

            If lastDay < 31 Then
                ws.Range(ws.Cells(r, COL_DAY_START + lastDay), ws.Cells(r, COL_DAY_END)).ClearContents
            End If
        End If
    Next r
End Sub

Private Sub RecalculateDetailTotals(ByVal ws As Worksheet, ByVal lastDay As Long)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long, c As Long, total As Double
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value))) = "DETAIL" Then
            total = 0
            For c = COL_DAY_START To COL_DAY_START + lastDay - 1
                total = total + NzNumber(ws.Cells(r, c).Value)
            Next c
            If total <> 0 Then ws.Cells(r, COL_MONTH_TOTAL).Value = total
        End If
    Next r
End Sub

Private Sub RecalculateSubtotals(ByVal ws As Worksheet, ByVal subtotalRows As Collection, ByVal lastDay As Long)
    Dim subtotalRow As Variant
    For Each subtotalRow In subtotalRows
        Dim groupName As String
        groupName = Trim$(CStr(ws.Cells(CLng(subtotalRow), COL_GROUP).Value))

        Dim c As Long
        For c = COL_DAY_START To COL_DAY_START + lastDay - 1
            Dim dayTotal As Double
            dayTotal = SumDetailByGroup(ws, groupName, c)
            If dayTotal <> 0 Then ws.Cells(CLng(subtotalRow), c).Value = dayTotal
        Next c

        Dim monthTotal As Double
        monthTotal = SumDetailByGroup(ws, groupName, COL_MONTH_TOTAL)
        If monthTotal <> 0 Then ws.Cells(CLng(subtotalRow), COL_MONTH_TOTAL).Value = monthTotal
    Next subtotalRow
End Sub

Private Function SumDetailByGroup(ByVal ws As Worksheet, ByVal groupName As String, ByVal valueCol As Long) As Double
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    Dim r As Long, total As Double
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
    If IsNumeric(v) Then NzNumber = CDbl(v) Else NzNumber = 0
End Function

Private Sub InitLog(ByVal ws As Worksheet)
    ws.Cells.ClearContents

    ws.Range("A1:J1").Value = Array( _
        "記録日時", "重要度", "分類", "ファイル", "行番号", _
        "日付", "支店コード", "業務コード", "値", "内容" _
    )
End Sub

Private Sub LogIssue(ByVal ws As Worksheet, ByVal severity As String, ByVal category As String, _
                     ByVal filePath As String, ByVal rowNumber As Long, ByVal dateKey As String, _
                     ByVal branchCode As String, ByVal bizCode As String, ByVal rawValue As Variant, _
                     ByVal message As String)
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
    ws.Cells(r, 9).Value = rawValue
    ws.Cells(r, 10).Value = message
End Sub
```
