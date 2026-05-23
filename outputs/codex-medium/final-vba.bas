Option Explicit

Private Const SHEET_MONTHLY As String = "月次集計"
Private Const SHEET_MASTER As String = "マスタ"
Private Const SHEET_LOG As String = "異常系"
Private Const SHEET_DAILY As String = "日次集計"

Private Const ROW_TYPE_DETAIL As String = "DETAIL"
Private Const ROW_TYPE_SUBTOTAL As String = "SUBTOTAL"

Private Const COL_ROW_TYPE As Long = 1
Private Const COL_SUBTOTAL_GROUP As Long = 2
Private Const COL_BRANCH_CODE As Long = 4
Private Const COL_BRANCH_NAME As Long = 5
Private Const COL_WORK_CODE As Long = 6
Private Const COL_WORK_NAME As Long = 7
Private Const COL_FIRST_DAY As Long = 8
Private Const COL_LAST_DAY As Long = 38
Private Const COL_MONTH_TOTAL As Long = 39

Private Const LOG_COL_COUNT As Long = 9

Public Sub ImportMonthlyClaimSummary()
    Dim wbMonthly As Workbook
    Dim wsMonthly As Worksheet
    Dim wsLog As Worksheet
    Dim targetYYMM As String
    Dim targetYear As Long
    Dim targetMonth As Long
    Dim daysInMonth As Long
    Dim dailyRoot As String
    Dim dailyFiles As Collection
    Dim validFilesByDate As Object
    Dim monthlyKeyToRow As Object
    Dim branchCodes As Object
    Dim workCodes As Object
    Dim dateKey As Variant
    Dim filePath As String
    Dim logRow As Long
    Dim importedFileCount As Long
    Dim importedRowCount As Long
    Dim skippedRowCount As Long

    On Error GoTo FatalError

    Set wbMonthly = ThisWorkbook

    If Len(wbMonthly.Path) = 0 Then
        MsgBox "月次ブックが未保存です。先に保存してから実行してください。" & vbCrLf & _
               "日次ファイル探索起点 ThisWorkbook.Path\daily\YYMM\ を決定できません。", vbCritical
        Exit Sub
    End If

    Set wsMonthly = GetRequiredSheet(wbMonthly, SHEET_MONTHLY)
    Set wsLog = GetRequiredSheet(wbMonthly, SHEET_LOG)

    targetYYMM = GetMonthlyYYMM(wbMonthly.Name)
    If Len(targetYYMM) = 0 Then
        Err.Raise vbObjectError + 1001, , _
            "月次ブック名が想定形式ではありません。形式: 月次クレーム集計YYMM.xlsm"
    End If

    targetYear = 2000 + CLng(Left$(targetYYMM, 2))
    targetMonth = CLng(Right$(targetYYMM, 2))
    If targetMonth < 1 Or targetMonth > 12 Then
        Err.Raise vbObjectError + 1002, , "月次ブック名のYYMMが不正です: " & targetYYMM
    End If

    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))
    dailyRoot = wbMonthly.Path & Application.PathSeparator & "daily" & _
                Application.PathSeparator & targetYYMM & Application.PathSeparator

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    PrepareLogSheet wsLog, logRow
    AddLog wsLog, logRow, "INFO", "", "", "", "", "", "", _
           "処理開始。対象年月=" & targetYYMM

    Set monthlyKeyToRow = CreateObject("Scripting.Dictionary")
    Set branchCodes = CreateObject("Scripting.Dictionary")
    Set workCodes = CreateObject("Scripting.Dictionary")

    BuildMonthlyIndex wsMonthly, monthlyKeyToRow, branchCodes, workCodes, wsLog, logRow
    LoadMasterCodes wbMonthly, branchCodes, workCodes, wsLog, logRow

    ClearDetailDailyValues wsMonthly

    If Len(Dir$(dailyRoot, vbDirectory)) = 0 Then
        AddLog wsLog, logRow, "ERROR", ToLogPath(wbMonthly.Path, dailyRoot), "", "", "", "", "", _
               "日次ファイル探索フォルダが存在しません。"
        GoTo Finish
    End If

    Set dailyFiles = New Collection
    CollectFilesRecursive dailyRoot, dailyFiles

    Set validFilesByDate = CreateObject("Scripting.Dictionary")
    ClassifyDailyFiles dailyFiles, targetYYMM, wbMonthly.Path, validFilesByDate, wsLog, logRow

    For Each dateKey In validFilesByDate.Keys
        filePath = CStr(validFilesByDate(CStr(dateKey)))
        ImportOneDailyBook filePath, CStr(dateKey), wbMonthly.Path, wsMonthly, monthlyKeyToRow, _
                           branchCodes, workCodes, wsLog, logRow, _
                           importedRowCount, skippedRowCount
        importedFileCount = importedFileCount + 1
    Next dateKey

Finish:
    RecalculateMonthlyTotals wsMonthly, daysInMonth
    RecalculateSubtotals wsMonthly, daysInMonth

    AddLog wsLog, logRow, "INFO", "", "", "", "", "", "", _
           "処理終了。取込ファイル数=" & importedFileCount & _
           " / 取込行数=" & importedRowCount & _
           " / スキップ行数=" & skippedRowCount

    wsLog.Columns("A:I").AutoFit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "月次クレーム集計の取込が完了しました。" & vbCrLf & _
           "詳細は「" & SHEET_LOG & "」シートを確認してください。", vbInformation
    Exit Sub

FatalError:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If Not wsLog Is Nothing Then
        AddLog wsLog, logRow, "FATAL", "", "", "", "", "", "", Err.Description
    End If

    MsgBox "処理を中断しました。" & vbCrLf & Err.Description, vbCritical
End Sub

Private Function GetRequiredSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error GoTo NotFound
    Set GetRequiredSheet = wb.Worksheets(sheetName)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 1101, , "必要なシートがありません: " & sheetName
End Function

Private Function GetMonthlyYYMM(ByVal bookName As String) As String
    If bookName Like "月次クレーム集計####.xlsm" Then
        GetMonthlyYYMM = Mid$(bookName, Len("月次クレーム集計") + 1, 4)
    Else
        GetMonthlyYYMM = vbNullString
    End If
End Function

Private Function GetDailyYYMMDD(ByVal fileNameOnly As String) As String
    If fileNameOnly Like "クレーム集計######.xlsx" Then
        GetDailyYYMMDD = Mid$(fileNameOnly, Len("クレーム集計") + 1, 6)
    Else
        GetDailyYYMMDD = vbNullString
    End If
End Function

Private Sub PrepareLogSheet(ByVal wsLog As Worksheet, ByRef logRow As Long)
    wsLog.Range(wsLog.Cells(1, 1), wsLog.Cells(wsLog.Rows.Count, LOG_COL_COUNT)).ClearContents

    wsLog.Cells(1, 1).Value = "日時"
    wsLog.Cells(1, 2).Value = "レベル"
    wsLog.Cells(1, 3).Value = "ファイル"
    wsLog.Cells(1, 4).Value = "日付"
    wsLog.Cells(1, 5).Value = "行"
    wsLog.Cells(1, 6).Value = "支店コード"
    wsLog.Cells(1, 7).Value = "業務コード"
    wsLog.Cells(1, 8).Value = "件数"
    wsLog.Cells(1, 9).Value = "内容"

    logRow = 2
End Sub

Private Sub AddLog(ByVal wsLog As Worksheet, ByRef logRow As Long, _
                   ByVal level As String, ByVal filePath As String, ByVal dateText As String, _
                   ByVal sourceRow As Variant, ByVal branchCode As Variant, _
                   ByVal workCode As Variant, ByVal claimAmount As Variant, _
                   ByVal message As String)
    wsLog.Cells(logRow, 1).Value = Now
    wsLog.Cells(logRow, 2).Value = level
    wsLog.Cells(logRow, 3).Value = filePath
    wsLog.Cells(logRow, 4).Value = dateText
    wsLog.Cells(logRow, 5).Value = sourceRow
    wsLog.Cells(logRow, 6).Value = branchCode
    wsLog.Cells(logRow, 7).Value = workCode
    wsLog.Cells(logRow, 8).Value = claimAmount
    wsLog.Cells(logRow, 9).Value = message
    logRow = logRow + 1
End Sub

Private Function ToLogPath(ByVal basePath As String, ByVal fullPath As String) As String
    Dim normalizedBase As String

    If Len(basePath) = 0 Or Len(fullPath) = 0 Then
        ToLogPath = fullPath
        Exit Function
    End If

    normalizedBase = basePath
    If Right$(normalizedBase, 1) <> Application.PathSeparator Then
        normalizedBase = normalizedBase & Application.PathSeparator
    End If

    If LCase$(Left$(fullPath, Len(normalizedBase))) = LCase$(normalizedBase) Then
        ToLogPath = Mid$(fullPath, Len(normalizedBase) + 1)
    Else
        ToLogPath = fullPath
    End If
End Function

Private Sub CollectFilesRecursive(ByVal folderPath As String, ByVal files As Collection)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim fileItem As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each fileItem In folder.Files
        files.Add fileItem.Path
    Next fileItem

    For Each subFolder In folder.SubFolders
        CollectFilesRecursive subFolder.Path, files
    Next subFolder
End Sub

Private Sub ClassifyDailyFiles(ByVal dailyFiles As Collection, ByVal targetYYMM As String, _
                               ByVal basePath As String, ByVal validFilesByDate As Object, _
                               ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim sortedFiles As Variant
    Dim i As Long
    Dim filePath As String
    Dim fileNameOnly As String
    Dim yymmdd As String
    Dim yymm As String
    Dim dayText As String
    Dim daysInFileMonth As Long

    sortedFiles = CollectionToSortedArray(dailyFiles)

    For i = LBound(sortedFiles) To UBound(sortedFiles)
        filePath = CStr(sortedFiles(i))
        fileNameOnly = Mid$(filePath, InStrRev(filePath, Application.PathSeparator) + 1)

        If LCase$(Right$(fileNameOnly, 5)) <> ".xlsx" Then
            AddLog wsLog, logRow, "WARN", ToLogPath(basePath, filePath), "", "", "", "", "", _
                   "xlsx以外のファイルのため処理対象外です。"
            GoTo ContinueNext
        End If

        yymmdd = GetDailyYYMMDD(fileNameOnly)
        If Len(yymmdd) = 0 Then
            AddLog wsLog, logRow, "ERROR", ToLogPath(basePath, filePath), "", "", "", "", "", _
                   "日次ファイル名が想定形式ではありません。形式: クレーム集計YYMMDD.xlsx"
            GoTo ContinueNext
        End If

        yymm = Left$(yymmdd, 4)
        dayText = Right$(yymmdd, 2)

        If yymm <> targetYYMM Then
            AddLog wsLog, logRow, "ERROR", ToLogPath(basePath, filePath), yymmdd, "", "", "", "", _
                   "対象月と異なる日次ファイルです。"
            GoTo ContinueNext
        End If

        daysInFileMonth = Day(DateSerial(2000 + CLng(Left$(yymm, 2)), CLng(Right$(yymm, 2)) + 1, 0))
        If CLng(dayText) < 1 Or CLng(dayText) > daysInFileMonth Then
            AddLog wsLog, logRow, "ERROR", ToLogPath(basePath, filePath), yymmdd, "", "", "", "", _
                   "日次ファイル名の日付が不正です。"
            GoTo ContinueNext
        End If

        If validFilesByDate.Exists(yymmdd) Then
            AddLog wsLog, logRow, "ERROR", ToLogPath(basePath, filePath), yymmdd, "", "", "", "", _
                   "同じ日付の重複日次ファイルです。フルパス昇順の先頭1件のみ処理し、このファイルは処理対象外にします。採用済み=" & _
                   ToLogPath(basePath, CStr(validFilesByDate(yymmdd)))
        Else
            validFilesByDate.Add yymmdd, filePath
        End If

ContinueNext:
    Next i
End Sub

Private Function CollectionToSortedArray(ByVal items As Collection) As Variant
    Dim arr() As String
    Dim i As Long
    Dim j As Long
    Dim temp As String

    If items.Count = 0 Then
        ReDim arr(0 To -1)
        CollectionToSortedArray = arr
        Exit Function
    End If

    ReDim arr(1 To items.Count)
    For i = 1 To items.Count
        arr(i) = CStr(items(i))
    Next i

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    CollectionToSortedArray = arr
End Function

Private Sub BuildMonthlyIndex(ByVal wsMonthly As Worksheet, ByVal monthlyKeyToRow As Object, _
                              ByVal branchCodes As Object, ByVal workCodes As Object, _
                              ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim lastRow As Long
    Dim r As Long
    Dim branchCode As String
    Dim workCode As String
    Dim key As String

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(wsMonthly.Cells(r, COL_ROW_TYPE).Value))) = ROW_TYPE_DETAIL Then
            branchCode = Trim$(CStr(wsMonthly.Cells(r, COL_BRANCH_CODE).Value))
            workCode = Trim$(CStr(wsMonthly.Cells(r, COL_WORK_CODE).Value))
            key = MakeKey(branchCode, workCode)

            If Len(branchCode) > 0 Then branchCodes(branchCode) = True
            If Len(workCode) > 0 Then workCodes(workCode) = True

            If monthlyKeyToRow.Exists(key) Then
                AddLog wsLog, logRow, "ERROR", "", "", r, branchCode, workCode, "", _
                       "月次集計に同じ支店コード+業務コードのDETAIL行が複数あります。先に見つかった行を使用します。"
            Else
                monthlyKeyToRow.Add key, r
            End If
        End If
    Next r
End Sub

Private Sub LoadMasterCodes(ByVal wb As Workbook, ByVal branchCodes As Object, ByVal workCodes As Object, _
                            ByVal wsLog As Worksheet, ByRef logRow As Long)
    Dim wsMaster As Worksheet
    Dim branchCodeCol As Long
    Dim workCodeCol As Long
    Dim mapBranchCol As Long
    Dim mapWorkCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim v As String
    Dim loadedBranch As Boolean
    Dim loadedWork As Boolean

    On Error Resume Next
    Set wsMaster = wb.Worksheets(SHEET_MASTER)
    On Error GoTo 0

    If wsMaster Is Nothing Then
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタシートがないため、月次集計のDETAIL行にあるコードを既知コードとして扱います。"
        Exit Sub
    End If

    branchCodeCol = FindHeaderColumnInRow(wsMaster, 2, "支店コード")
    If branchCodeCol > 0 Then
        For r = 3 To 11
            v = Trim$(CStr(wsMaster.Cells(r, branchCodeCol).Value))
            If Len(v) > 0 Then
                branchCodes(v) = True
                loadedBranch = True
            End If
        Next r
    Else
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタの2行目に支店マスタ見出し「支店コード」が見つかりません。月次DETAIL行の支店コードをフォールバックとして使用します。"
    End If

    workCodeCol = FindHeaderColumnInRow(wsMaster, 12, "業務コード")
    If workCodeCol > 0 Then
        For r = 13 To 23
            v = Trim$(CStr(wsMaster.Cells(r, workCodeCol).Value))
            If Len(v) > 0 Then
                workCodes(v) = True
                loadedWork = True
            End If
        Next r
    Else
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタの12行目に業務マスタ見出し「業務コード」が見つかりません。月次DETAIL行の業務コードをフォールバックとして使用します。"
    End If

    mapBranchCol = FindHeaderColumnInRow(wsMaster, 24, "支店コード")
    mapWorkCol = FindHeaderColumnInRow(wsMaster, 24, "業務コード")
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row

    If mapBranchCol > 0 Then
        For r = 25 To lastRow
            v = Trim$(CStr(wsMaster.Cells(r, mapBranchCol).Value))
            If Len(v) > 0 Then
                branchCodes(v) = True
                loadedBranch = True
            End If
        Next r
    Else
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタの24行目に支店・業務対応表見出し「支店コード」が見つかりません。"
    End If

    If mapWorkCol > 0 Then
        For r = 25 To lastRow
            v = Trim$(CStr(wsMaster.Cells(r, mapWorkCol).Value))
            If Len(v) > 0 Then
                workCodes(v) = True
                loadedWork = True
            End If
        Next r
    Else
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタの24行目に支店・業務対応表見出し「業務コード」が見つかりません。"
    End If

    If Not loadedBranch Then
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタから支店コードを読み取れませんでした。月次DETAIL行の支店コードのみで判定します。"
    End If

    If Not loadedWork Then
        AddLog wsLog, logRow, "WARN", "", "", "", "", "", "", _
               "マスタから業務コードを読み取れませんでした。月次DETAIL行の業務コードのみで判定します。"
    End If
End Sub

Private Function FindHeaderColumnInRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(headerRow, c).Value)) = headerText Then
            FindHeaderColumnInRow = c
            Exit Function
        End If
    Next c

    FindHeaderColumnInRow = 0
End Function

Private Sub ClearDetailDailyValues(ByVal wsMonthly As Worksheet)
    Dim lastRow As Long
    Dim r As Long

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(wsMonthly.Cells(r, COL_ROW_TYPE).Value))) = ROW_TYPE_DETAIL Then
            wsMonthly.Range(wsMonthly.Cells(r, COL_FIRST_DAY), wsMonthly.Cells(r, COL_LAST_DAY)).ClearContents
        End If
    Next r
End Sub

Private Sub ImportOneDailyBook(ByVal filePath As String, ByVal yymmdd As String, ByVal basePath As String, _
                               ByVal wsMonthly As Worksheet, ByVal monthlyKeyToRow As Object, _
                               ByVal branchCodes As Object, ByVal workCodes As Object, _
                               ByVal wsLog As Worksheet, ByRef logRow As Long, _
                               ByRef importedRowCount As Long, ByRef skippedRowCount As Long)
    Dim wbDaily As Workbook
    Dim wsDaily As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim dayNumber As Long
    Dim targetCol As Long
    Dim processType As String
    Dim branchCode As String
    Dim workCode As String
    Dim claimValue As Variant
    Dim key As String
    Dim targetRow As Long
    Dim logPath As String

    On Error GoTo OpenError

    logPath = ToLogPath(basePath, filePath)
    Set wbDaily = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)

    On Error Resume Next
    Set wsDaily = wbDaily.Worksheets(SHEET_DAILY)
    On Error GoTo OpenError

    If wsDaily Is Nothing Then
        AddLog wsLog, logRow, "ERROR", logPath, yymmdd, "", "", "", "", _
               "日次集計シートがありません。"
        skippedRowCount = skippedRowCount + 1
        GoTo CloseDaily
    End If

    dayNumber = CLng(Right$(yymmdd, 2))
    targetCol = COL_FIRST_DAY + dayNumber - 1

    lastRow = wsDaily.Cells(wsDaily.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        processType = LCase$(Trim$(CStr(wsDaily.Cells(r, 1).Value)))
        branchCode = Trim$(CStr(wsDaily.Cells(r, 2).Value))
        workCode = Trim$(CStr(wsDaily.Cells(r, 4).Value))
        claimValue = wsDaily.Cells(r, 6).Value

        If processType <> "ok" Then
            AddLog wsLog, logRow, "WARN", logPath, yymmdd, r, branchCode, workCode, claimValue, _
                   "処理区分がokではないため転記しません。処理区分=" & processType
            skippedRowCount = skippedRowCount + 1
            GoTo ContinueRow
        End If

        If Not branchCodes.Exists(branchCode) Then
            AddLog wsLog, logRow, "ERROR", logPath, yymmdd, r, branchCode, workCode, claimValue, _
                   "未知の支店コードです。"
            skippedRowCount = skippedRowCount + 1
            GoTo ContinueRow
        End If

        If Not workCodes.Exists(workCode) Then
            AddLog wsLog, logRow, "ERROR", logPath, yymmdd, r, branchCode, workCode, claimValue, _
                   "未知の業務コードです。"
            skippedRowCount = skippedRowCount + 1
            GoTo ContinueRow
        End If

        If Len(Trim$(CStr(claimValue))) = 0 Then
            AddLog wsLog, logRow, "ERROR", logPath, yymmdd, r, branchCode, workCode, claimValue, _
                   "クレーム件数が空欄です。"
            skippedRowCount = skippedRowCount + 1
            GoTo ContinueRow
        End If

        If Not IsNumeric(claimValue) Then
            AddLog wsLog, logRow, "ERROR", logPath, yymmdd, r, branchCode, workCode, claimValue, _
                   "クレーム件数が数値ではありません。"
            skippedRowCount = skippedRowCount + 1
            GoTo ContinueRow
        End If

        key = MakeKey(branchCode, workCode)

        If Not monthlyKeyToRow.Exists(key) Then
            AddLog wsLog, logRow, "ERROR", logPath, yymmdd, r, branchCode, workCode, claimValue, _
                   "月次集計に存在しない支店コード+業務コードです。"
            skippedRowCount = skippedRowCount + 1
            GoTo ContinueRow
        End If

        targetRow = CLng(monthlyKeyToRow(key))
        wsMonthly.Cells(targetRow, targetCol).Value = NzNumber(wsMonthly.Cells(targetRow, targetCol).Value) + CDbl(claimValue)
        importedRowCount = importedRowCount + 1

ContinueRow:
    Next r

CloseDaily:
    wbDaily.Close SaveChanges:=False
    Exit Sub

OpenError:
    AddLog wsLog, logRow, "ERROR", ToLogPath(basePath, filePath), yymmdd, "", "", "", "", _
           "日次ブックを開く、または読み込む際にエラーが発生しました: " & Err.Description
    On Error Resume Next
    If Not wbDaily Is Nothing Then wbDaily.Close SaveChanges:=False
End Sub

Private Sub RecalculateMonthlyTotals(ByVal wsMonthly As Worksheet, ByVal daysInMonth As Long)
    Dim lastRow As Long
    Dim r As Long
    Dim c As Long
    Dim total As Double

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(wsMonthly.Cells(r, COL_ROW_TYPE).Value))) = ROW_TYPE_DETAIL Then
            total = 0

            For c = COL_FIRST_DAY To COL_FIRST_DAY + daysInMonth - 1
                total = total + NzNumber(wsMonthly.Cells(r, c).Value)
            Next c

            wsMonthly.Cells(r, COL_MONTH_TOTAL).ClearContents
            If total <> 0 Then wsMonthly.Cells(r, COL_MONTH_TOTAL).Value = total
        End If
    Next r
End Sub

Private Sub RecalculateSubtotals(ByVal wsMonthly As Worksheet, ByVal daysInMonth As Long)
    Dim lastRow As Long
    Dim r As Long
    Dim detailRow As Long
    Dim c As Long
    Dim groupName As String
    Dim total As Double

    lastRow = wsMonthly.Cells(wsMonthly.Rows.Count, COL_ROW_TYPE).End(xlUp).Row

    For r = 2 To lastRow
        If UCase$(Trim$(CStr(wsMonthly.Cells(r, COL_ROW_TYPE).Value))) = ROW_TYPE_SUBTOTAL Then
            groupName = Trim$(CStr(wsMonthly.Cells(r, COL_SUBTOTAL_GROUP).Value))

            wsMonthly.Range(wsMonthly.Cells(r, COL_FIRST_DAY), wsMonthly.Cells(r, COL_LAST_DAY)).ClearContents
            wsMonthly.Cells(r, COL_MONTH_TOTAL).ClearContents

            For c = COL_FIRST_DAY To COL_FIRST_DAY + daysInMonth - 1
                total = 0

                For detailRow = 2 To lastRow
                    If UCase$(Trim$(CStr(wsMonthly.Cells(detailRow, COL_ROW_TYPE).Value))) = ROW_TYPE_DETAIL Then
                        If Trim$(CStr(wsMonthly.Cells(detailRow, COL_SUBTOTAL_GROUP).Value)) = groupName Then
                            total = total + NzNumber(wsMonthly.Cells(detailRow, c).Value)
                        End If
                    End If
                Next detailRow

                If total <> 0 Then wsMonthly.Cells(r, c).Value = total
            Next c

            total = 0
            For c = COL_FIRST_DAY To COL_FIRST_DAY + daysInMonth - 1
                total = total + NzNumber(wsMonthly.Cells(r, c).Value)
            Next c

            If total <> 0 Then wsMonthly.Cells(r, COL_MONTH_TOTAL).Value = total
        End If
    Next r
End Sub

Private Function MakeKey(ByVal branchCode As String, ByVal workCode As String) As String
    MakeKey = Trim$(branchCode) & "|" & Trim$(workCode)
End Function

Private Function NzNumber(ByVal value As Variant) As Double
    If IsNumeric(value) Then
        NzNumber = CDbl(value)
    Else
        NzNumber = 0
    End If
End Function
