# codex-multi-agent-review 最終出力

```vba
Option Explicit

Private Const SHEET_MONTHLY As String = "月次集計"
Private Const SHEET_MASTER As String = "マスタ"
Private Const SHEET_LOG As String = "異常系"
Private Const SHEET_DAILY As String = "日次集計"

Private Const MONTHLY_PREFIX As String = "月次クレーム集計"
Private Const MONTHLY_SUFFIX As String = ".xlsm"
Private Const DAILY_PREFIX As String = "クレーム集計"
Private Const DAILY_SUFFIX As String = ".xlsx"

Private Const ROW_HEADER As Long = 1
Private Const ROW_DATA_START As Long = 2

Private Const COL_ROW_TYPE As Long = 1
Private Const COL_SUBTOTAL_GROUP As Long = 2
Private Const COL_BRANCH_CODE As Long = 4
Private Const COL_BUSINESS_CODE As Long = 6
Private Const COL_FIRST_DAY As Long = 8
Private Const COL_LAST_DAY As Long = 38
Private Const COL_MONTH_TOTAL As Long = 39

Private Const DAILY_COL_STATUS As Long = 1
Private Const DAILY_COL_BRANCH_CODE As Long = 2
Private Const DAILY_COL_BUSINESS_CODE As Long = 4
Private Const DAILY_COL_COUNT As Long = 6

Private Const KEY_SEP As String = "|"

Private mLogWs As Worksheet
Private mLogRow As Long
Private mLogCount As Long

Public Sub ImportMonthlyClaimCounts()
    Dim wb As Workbook
    Dim wsMonthly As Worksheet
    Dim wsMaster As Worksheet
    Dim wsLog As Worksheet
    Dim targetYYMM As String
    Dim targetYear As Long
    Dim targetMonth As Long
    Dim daysInMonth As Long
    Dim dailyRoot As String
    Dim detailRows As Object
    Dim groupDetails As Object
    Dim subtotalRows As Object
    Dim masterBranches As Object
    Dim masterBusinesses As Object
    Dim processedDates As Object
    Dim fso As Object
    Dim filePaths As Collection
    Dim sortedPaths() As String
    Dim i As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim started As Date
    Dim aborted As Boolean
    Dim fatalMessage As String

    Set wb = ThisWorkbook

    If Len(wb.Path) = 0 Then
        MsgBox "月次ブックが未保存です。先に 月次クレーム集計YYMM.xlsm として保存してから実行してください。", vbExclamation
        Exit Sub
    End If

    On Error GoTo FatalError

    started = Now

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wsMonthly = GetRequiredWorksheet(wb, SHEET_MONTHLY)
    Set wsMaster = GetRequiredWorksheet(wb, SHEET_MASTER)
    Set wsLog = GetRequiredWorksheet(wb, SHEET_LOG)
    Set mLogWs = wsLog
    PrepareLogSheet wsLog

    If Not TryParseMonthlyFileName(wb.Name, targetYYMM, targetYear, targetMonth) Then
        AddLog "月次ファイル名形式違い", ToLogPath(wb.FullName), "", 0, "", "", "", "", _
               "月次ブック名は 月次クレーム集計YYMM.xlsm 形式にしてください。"
        Err.Raise vbObjectError + 1000, , "月次ブック名が想定形式ではありません。"
    End If

    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))
    dailyRoot = wb.Path & Application.PathSeparator & "daily" & Application.PathSeparator & targetYYMM

    Set detailRows = CreateTextDictionary()
    Set groupDetails = CreateTextDictionary()
    Set subtotalRows = CreateTextDictionary()
    Set masterBranches = CreateTextDictionary()
    Set masterBusinesses = CreateTextDictionary()
    Set processedDates = CreateTextDictionary()
    Set fso = CreateObject("Scripting.FileSystemObject")

    BuildMonthlyIndex wsMonthly, detailRows, groupDetails, subtotalRows
    BuildMasterSets wsMaster, wsMonthly, masterBranches, masterBusinesses
    ResetMonthlyDetailDays wsMonthly

    If Not fso.FolderExists(dailyRoot) Then
        AddLog "日次フォルダなし", ToLogPath(dailyRoot), "", 0, "", "", "", "", _
               "探索起点フォルダが存在しないため、日次ブックの読み込みをスキップしました。"
    Else
        Set filePaths = New Collection
        CollectDailyFilePaths fso.GetFolder(dailyRoot), filePaths

        If filePaths.Count > 0 Then
            sortedPaths = CollectionToSortedArray(filePaths)
            For i = LBound(sortedPaths) To UBound(sortedPaths)
                ProcessDailyFile sortedPaths(i), fso.GetFileName(sortedPaths(i)), targetYYMM, _
                                 detailRows, masterBranches, masterBusinesses, processedDates
            Next i
        End If
    End If

    RecalculateMonthlyTotals wsMonthly, groupDetails, subtotalRows, daysInMonth

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    On Error GoTo 0

    If aborted Then
        MsgBox "月次クレーム集計を中断しました。" & vbCrLf & _
               "理由: " & fatalMessage & vbCrLf & _
               "ログ件数: " & CStr(mLogCount), vbExclamation
    Else
        MsgBox "月次クレーム集計が完了しました。" & vbCrLf & _
               "対象月: " & targetYYMM & vbCrLf & _
               "ログ件数: " & CStr(mLogCount) & vbCrLf & _
               "開始: " & Format$(started, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
               "終了: " & Format$(Now, "yyyy/mm/dd hh:nn:ss"), vbInformation
    End If
    Exit Sub

FatalError:
    aborted = True
    fatalMessage = Err.Number & " " & Err.Description
    AddLog "致命的エラー", ToLogPath(ThisWorkbook.FullName), "", 0, "", "", "", "", _
           "処理を中断しました。" & fatalMessage
    Resume CleanExit
End Sub

Private Function GetRequiredWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetRequiredWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetRequiredWorksheet Is Nothing Then
        Err.Raise vbObjectError + 1001, , "必要なシートがありません: " & sheetName
    End If
End Function

Private Sub PrepareLogSheet(ByVal ws As Worksheet)
    ws.Range("A:J").ClearContents
    ws.Range("A1:J1").Value = Array("発生日時", "種別", "ファイルパス", "シート名", "行番号", _
                                    "支店コード", "業務コード", "日付", "クレーム件数元値", "内容")
    mLogRow = 2
    mLogCount = 0
End Sub

Private Sub AddLog(ByVal logType As String, ByVal filePath As String, ByVal sheetName As String, _
                   ByVal rowNumber As Long, ByVal branchCode As String, ByVal businessCode As String, _
                   ByVal targetDate As String, ByVal claimCountRaw As String, ByVal message As String)
    If mLogWs Is Nothing Then Exit Sub

    With mLogWs
        .Cells(mLogRow, 1).Value = Now
        .Cells(mLogRow, 2).Value = logType
        .Cells(mLogRow, 3).Value = filePath
        .Cells(mLogRow, 4).Value = sheetName
        If rowNumber > 0 Then .Cells(mLogRow, 5).Value = rowNumber
        .Cells(mLogRow, 6).Value = branchCode
        .Cells(mLogRow, 7).Value = businessCode
        .Cells(mLogRow, 8).Value = targetDate
        .Cells(mLogRow, 9).Value = claimCountRaw
        .Cells(mLogRow, 10).Value = message
    End With

    mLogRow = mLogRow + 1
    mLogCount = mLogCount + 1
End Sub

Private Function ToLogPath(ByVal fullPath As String) As String
    Dim basePath As String
    Dim normalizedBase As String
    Dim normalizedFull As String

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Or Len(fullPath) = 0 Then
        ToLogPath = fullPath
        Exit Function
    End If

    normalizedBase = Replace(basePath, "/", "\")
    normalizedFull = Replace(fullPath, "/", "\")

    If LCase$(normalizedFull) = LCase$(normalizedBase) Then
        ToLogPath = "."
    ElseIf LCase$(Left$(normalizedFull, Len(normalizedBase) + 1)) = LCase$(normalizedBase & "\") Then
        ToLogPath = Mid$(normalizedFull, Len(normalizedBase) + 2)
    Else
        ToLogPath = fullPath
    End If
End Function

Private Function TryParseMonthlyFileName(ByVal fileName As String, ByRef yymm As String, _
                                         ByRef yyyy As Long, ByRef mm As Long) As Boolean
    Dim core As String

    If LCase$(Right$(fileName, Len(MONTHLY_SUFFIX))) <> MONTHLY_SUFFIX Then Exit Function
    If Left$(fileName, Len(MONTHLY_PREFIX)) <> MONTHLY_PREFIX Then Exit Function
    If Len(fileName) <> Len(MONTHLY_PREFIX) + 4 + Len(MONTHLY_SUFFIX) Then Exit Function

    core = Mid$(fileName, Len(MONTHLY_PREFIX) + 1, 4)
    If Not IsAllDigits(core) Then Exit Function

    yymm = core
    yyyy = 2000 + CLng(Left$(core, 2))
    mm = CLng(Right$(core, 2))
    If mm < 1 Or mm > 12 Then Exit Function

    TryParseMonthlyFileName = True
End Function

Private Function TryParseDailyFileName(ByVal fileName As String, ByRef yymmdd As String, _
                                       ByRef yymm As String, ByRef yyyy As Long, _
                                       ByRef mm As Long, ByRef dd As Long) As Boolean
    Dim core As String
    Dim d As Date

    If LCase$(Right$(fileName, Len(DAILY_SUFFIX))) <> DAILY_SUFFIX Then Exit Function
    If Left$(fileName, Len(DAILY_PREFIX)) <> DAILY_PREFIX Then Exit Function
    If Len(fileName) <> Len(DAILY_PREFIX) + 6 + Len(DAILY_SUFFIX) Then Exit Function

    core = Mid$(fileName, Len(DAILY_PREFIX) + 1, 6)
    If Not IsAllDigits(core) Then Exit Function

    yymmdd = core
    yymm = Left$(core, 4)
    yyyy = 2000 + CLng(Left$(core, 2))
    mm = CLng(Mid$(core, 3, 2))
    dd = CLng(Right$(core, 2))

    If mm < 1 Or mm > 12 Then Exit Function

    On Error Resume Next
    d = DateSerial(yyyy, mm, dd)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If Year(d) <> yyyy Or Month(d) <> mm Or Day(d) <> dd Then Exit Function

    TryParseDailyFileName = True
End Function

Private Function IsAllDigits(ByVal text As String) As Boolean
    Dim i As Long

    If Len(text) = 0 Then Exit Function

    For i = 1 To Len(text)
        If Mid$(text, i, 1) < "0" Or Mid$(text, i, 1) > "9" Then Exit Function
    Next i

    IsAllDigits = True
End Function

Private Function CreateTextDictionary() As Object
    Set CreateTextDictionary = CreateObject("Scripting.Dictionary")
    CreateTextDictionary.CompareMode = vbTextCompare
End Function

Private Sub BuildMasterSets(ByVal wsMaster As Worksheet, ByVal wsMonthly As Worksheet, _
                            ByVal branches As Object, ByVal businesses As Object)
    Dim missingHeader As Boolean

    ReadSingleCodeMasterSection wsMaster, "支店マスタ", 2, 3, 11, "支店コード", branches, missingHeader
    ReadSingleCodeMasterSection wsMaster, "業務マスタ", 12, 13, 23, "業務コード", businesses, missingHeader
    ReadBranchBusinessMapSection wsMaster, 24, 25, branches, businesses, missingHeader

    If missingHeader Or branches.Count = 0 Or businesses.Count = 0 Then
        AddLog "マスタフォールバック", ToLogPath(ThisWorkbook.FullName), wsMaster.Name, 0, "", "", "", "", _
               "マスタの一部見出しが見つからない、またはコードを取得できないため、月次DETAIL行から支店コード・業務コードを補完します。"
        AddMasterCodesFromMonthlyDetail wsMonthly, branches, businesses
    End If
End Sub

Private Sub ReadSingleCodeMasterSection(ByVal ws As Worksheet, ByVal sectionName As String, _
                                        ByVal headerRow As Long, ByVal firstDataRow As Long, _
                                        ByVal lastDataRow As Long, ByVal headerText As String, _
                                        ByVal codes As Object, ByRef missingHeader As Boolean)
    Dim codeCol As Long
    Dim r As Long
    Dim codeValue As String

    codeCol = FindHeaderColumnInRow(ws, headerRow, headerText)

    If codeCol = 0 Then
        missingHeader = True
        AddLog "マスタ見出し不足", ToLogPath(ThisWorkbook.FullName), ws.Name, headerRow, "", "", "", "", _
               sectionName & " の見出し行に " & headerText & " が見つかりません。"
        Exit Sub
    End If

    For r = firstDataRow To lastDataRow
        codeValue = NormalizeCode(ws.Cells(r, codeCol).Value)
        If Len(codeValue) > 0 Then
            If Not codes.Exists(codeValue) Then codes.Add codeValue, True
        End If
    Next r
End Sub

Private Sub ReadBranchBusinessMapSection(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                         ByVal firstDataRow As Long, ByVal branches As Object, _
                                         ByVal businesses As Object, ByRef missingHeader As Boolean)
    Dim branchCol As Long
    Dim businessCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim branchCode As String
    Dim businessCode As String

    branchCol = FindHeaderColumnInRow(ws, headerRow, "支店コード")
    businessCol = FindHeaderColumnInRow(ws, headerRow, "業務コード")

    If branchCol = 0 Or businessCol = 0 Then
        missingHeader = True
        AddLog "マスタ見出し不足", ToLogPath(ThisWorkbook.FullName), ws.Name, headerRow, "", "", "", "", _
               "支店・業務対応表の見出し行に 支店コード または 業務コード が見つかりません。"
        Exit Sub
    End If

    lastRow = MaxLong(LastRowInColumn(ws, branchCol), LastRowInColumn(ws, businessCol))

    For r = firstDataRow To lastRow
        branchCode = NormalizeCode(ws.Cells(r, branchCol).Value)
        businessCode = NormalizeCode(ws.Cells(r, businessCol).Value)

        If Len(branchCode) > 0 Then
            If Not branches.Exists(branchCode) Then branches.Add branchCode, True
        End If

        If Len(businessCode) > 0 Then
            If Not businesses.Exists(businessCode) Then businesses.Add businessCode, True
        End If
    Next r
End Sub

Private Sub AddMasterCodesFromMonthlyDetail(ByVal ws As Worksheet, ByVal branches As Object, ByVal businesses As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim rowType As String
    Dim branchCode As String
    Dim businessCode As String

    lastRow = GetLastUsedRow(ws, COL_ROW_TYPE, COL_MONTH_TOTAL)

    For r = ROW_DATA_START To lastRow
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))

        If rowType = "DETAIL" Then
            branchCode = NormalizeCode(ws.Cells(r, COL_BRANCH_CODE).Value)
            businessCode = NormalizeCode(ws.Cells(r, COL_BUSINESS_CODE).Value)

            If Len(branchCode) > 0 Then
                If Not branches.Exists(branchCode) Then branches.Add branchCode, True
            End If

            If Len(businessCode) > 0 Then
                If Not businesses.Exists(businessCode) Then businesses.Add businessCode, True
            End If
        End If
    Next r
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
End Function

Private Sub BuildMonthlyIndex(ByVal ws As Worksheet, ByVal detailRows As Object, _
                              ByVal groupDetails As Object, ByVal subtotalRows As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim rowType As String
    Dim branchCode As String
    Dim businessCode As String
    Dim groupName As String
    Dim key As String

    lastRow = GetLastUsedRow(ws, COL_ROW_TYPE, COL_MONTH_TOTAL)

    For r = ROW_DATA_START To lastRow
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))
        groupName = Trim$(CStr(ws.Cells(r, COL_SUBTOTAL_GROUP).Value))

        Select Case rowType
            Case "DETAIL"
                branchCode = NormalizeCode(ws.Cells(r, COL_BRANCH_CODE).Value)
                businessCode = NormalizeCode(ws.Cells(r, COL_BUSINESS_CODE).Value)

                If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
                    AddLog "月次DETAILキー不備", ToLogPath(ThisWorkbook.FullName), ws.Name, r, branchCode, businessCode, "", "", _
                           "DETAIL行の支店コードまたは業務コードが空欄です。"
                Else
                    key = MakeKey(branchCode, businessCode)
                    If detailRows.Exists(key) Then
                        AddLog "月次DETAIL重複", ToLogPath(ThisWorkbook.FullName), ws.Name, r, branchCode, businessCode, "", "", _
                               "同じ支店コード+業務コードのDETAIL行が複数あります。先に見つかった行へ転記します。"
                    Else
                        detailRows.Add key, r
                    End If
                    AddRowToGroup groupDetails, groupName, r
                End If

            Case "SUBTOTAL"
                If Len(groupName) = 0 Then
                    AddLog "月次SUBTOTALグループ不備", ToLogPath(ThisWorkbook.FullName), ws.Name, r, "", "", "", "", _
                           "SUBTOTAL行の小計グループが空欄です。"
                Else
                    AddRowToGroup subtotalRows, groupName, r
                End If
        End Select
    Next r
End Sub

Private Sub AddRowToGroup(ByVal dict As Object, ByVal groupName As String, ByVal rowNumber As Long)
    Dim rows As Collection

    If Len(groupName) = 0 Then Exit Sub

    If dict.Exists(groupName) Then
        Set rows = dict(groupName)
    Else
        Set rows = New Collection
        dict.Add groupName, rows
    End If

    rows.Add rowNumber
End Sub

Private Sub ResetMonthlyDetailDays(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim rowType As String

    lastRow = GetLastUsedRow(ws, COL_ROW_TYPE, COL_MONTH_TOTAL)

    For r = ROW_DATA_START To lastRow
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))
        If rowType = "DETAIL" Then
            ws.Range(ws.Cells(r, COL_FIRST_DAY), ws.Cells(r, COL_LAST_DAY)).ClearContents
        End If
    Next r
End Sub

Private Sub CollectDailyFilePaths(ByVal folder As Object, ByVal filePaths As Collection)
    Dim file As Object
    Dim subFolder As Object
    Dim fileName As String

    For Each file In folder.Files
        fileName = CStr(file.Name)
        If Left$(fileName, 2) <> "~$" Then filePaths.Add CStr(file.Path)
    Next file

    For Each subFolder In folder.SubFolders
        CollectDailyFilePaths subFolder, filePaths
    Next subFolder
End Sub

Private Function CollectionToSortedArray(ByVal items As Collection) As String()
    Dim arr() As String
    Dim i As Long

    ReDim arr(1 To items.Count)

    For i = 1 To items.Count
        arr(i) = CStr(items(i))
    Next i

    SortStringArray arr, LBound(arr), UBound(arr)
    CollectionToSortedArray = arr
End Function

Private Sub SortStringArray(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
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

    If first < high Then SortStringArray arr, first, high
    If low < last Then SortStringArray arr, low, last
End Sub

Private Sub ProcessDailyFile(ByVal filePath As String, ByVal fileName As String, ByVal targetYYMM As String, _
                             ByVal detailRows As Object, ByVal masterBranches As Object, _
                             ByVal masterBusinesses As Object, ByVal processedDates As Object)
    Dim yymmdd As String
    Dim fileYYMM As String
    Dim fileYear As Long
    Dim fileMonth As Long
    Dim fileDay As Long
    Dim targetDateText As String
    Dim dailyWb As Workbook
    Dim dailyWs As Worksheet

    On Error GoTo FileError

    If Not TryParseDailyFileName(fileName, yymmdd, fileYYMM, fileYear, fileMonth, fileDay) Then
        AddLog "日次ファイル名形式違い", ToLogPath(filePath), "", 0, "", "", "", "", _
               "日次ブック名は クレーム集計YYMMDD.xlsx 形式にしてください。"
        Exit Sub
    End If

    If fileYYMM <> targetYYMM Then
        AddLog "対象月違い", ToLogPath(filePath), "", 0, "", "", yymmdd, "", _
               "月次対象月 " & targetYYMM & " と日次ファイルの年月 " & fileYYMM & " が一致しません。"
        Exit Sub
    End If

    targetDateText = Format$(DateSerial(fileYear, fileMonth, fileDay), "yyyy/mm/dd")

    If processedDates.Exists(yymmdd) Then
        AddLog "同じ日付の重複日次ファイル", ToLogPath(filePath), "", 0, "", "", targetDateText, "", _
               "同じ日付のファイルを既に処理済みです。採用済みファイル: " & CStr(processedDates(yymmdd))
        Exit Sub
    End If

    Set dailyWb = Workbooks.Open(Filename:=filePath, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)

    If Not WorksheetExists(dailyWb, SHEET_DAILY) Then
        AddLog "日次集計シートなし", ToLogPath(filePath), "", 0, "", "", targetDateText, "", _
               "日次ブックに 日次集計 シートがありません。"
        GoTo FileExit
    End If

    processedDates.Add yymmdd, ToLogPath(filePath)

    Set dailyWs = dailyWb.Worksheets(SHEET_DAILY)
    ProcessDailyRows dailyWs, filePath, targetDateText, fileDay, detailRows, masterBranches, masterBusinesses

FileExit:
    On Error Resume Next
    If Not dailyWb Is Nothing Then dailyWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

FileError:
    If Len(targetDateText) = 0 Then targetDateText = yymmdd
    AddLog "日次ブック処理エラー", ToLogPath(filePath), "", 0, "", "", targetDateText, "", _
           "日次ブックの処理中にエラーが発生しました。" & Err.Number & " " & Err.Description
    Resume FileExit
End Sub

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub ProcessDailyRows(ByVal ws As Worksheet, ByVal filePath As String, ByVal targetDateText As String, _
                             ByVal fileDay As Long, ByVal detailRows As Object, _
                             ByVal masterBranches As Object, ByVal masterBusinesses As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim statusText As String
    Dim branchCode As String
    Dim businessCode As String
    Dim countValue As Variant
    Dim countText As String
    Dim claimCount As Double
    Dim key As String
    Dim targetRow As Long
    Dim targetCol As Long
    Dim branchOk As Boolean
    Dim businessOk As Boolean
    Dim newValue As Double

    lastRow = GetLastUsedRow(ws, DAILY_COL_STATUS, DAILY_COL_COUNT)

    For r = ROW_DATA_START To lastRow
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))) = 0 Then GoTo ContinueRow

        statusText = LCase$(Trim$(CStr(ws.Cells(r, DAILY_COL_STATUS).Value)))
        branchCode = NormalizeCode(ws.Cells(r, DAILY_COL_BRANCH_CODE).Value)
        businessCode = NormalizeCode(ws.Cells(r, DAILY_COL_BUSINESS_CODE).Value)
        countValue = ws.Cells(r, DAILY_COL_COUNT).Value
        countText = ValueForLog(countValue)

        If statusText <> "ok" Then
            AddLog "処理区分対象外", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "処理区分が ok ではないため転記しません。値: " & ValueForLog(ws.Cells(r, DAILY_COL_STATUS).Value)
            GoTo ContinueRow
        End If

        If IsError(countValue) Then
            AddLog "クレーム件数非数値", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "クレーム件数がエラー値です。"
            GoTo ContinueRow
        End If

        If Len(Trim$(countText)) = 0 Then
            AddLog "クレーム件数空欄", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "クレーム件数が空欄です。"
            GoTo ContinueRow
        End If

        If Not IsNumeric(countText) Then
            AddLog "クレーム件数非数値", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "クレーム件数が数値ではありません。"
            GoTo ContinueRow
        End If

        claimCount = CDbl(countText)

        branchOk = masterBranches.Exists(branchCode)
        businessOk = masterBusinesses.Exists(businessCode)

        If Not branchOk Then
            AddLog "未知の支店コード", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "マスタに存在しない支店コードです。"
        End If

        If Not businessOk Then
            AddLog "未知の業務コード", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "マスタに存在しない業務コードです。"
        End If

        If Not branchOk Or Not businessOk Then GoTo ContinueRow

        key = MakeKey(branchCode, businessCode)

        If Not detailRows.Exists(key) Then
            AddLog "月次表に存在しないキー", ToLogPath(filePath), ws.Name, r, branchCode, businessCode, targetDateText, countText, _
                   "月次集計のDETAIL行に支店コード+業務コードが存在しません。"
            GoTo ContinueRow
        End If

        targetRow = CLng(detailRows(key))
        targetCol = COL_FIRST_DAY + fileDay - 1

        With ThisWorkbook.Worksheets(SHEET_MONTHLY).Cells(targetRow, targetCol)
            newValue = CDbl(Val(.Value)) + claimCount
            If newValue = 0 Then
                .ClearContents
            Else
                .Value = newValue
            End If
        End With

ContinueRow:
    Next r
End Sub

Private Sub RecalculateMonthlyTotals(ByVal ws As Worksheet, ByVal groupDetails As Object, _
                                     ByVal subtotalRows As Object, ByVal daysInMonth As Long)
    Dim lastRow As Long
    Dim r As Long
    Dim rowType As String
    Dim dayIndex As Long
    Dim subtotalGroup As Variant
    Dim subtotalList As Collection
    Dim detailList As Collection
    Dim subtotalRow As Variant
    Dim detailRow As Variant
    Dim total As Double
    Dim dayTotal As Double

    lastRow = GetLastUsedRow(ws, COL_ROW_TYPE, COL_MONTH_TOTAL)

    For r = ROW_DATA_START To lastRow
        rowType = UCase$(Trim$(CStr(ws.Cells(r, COL_ROW_TYPE).Value)))

        If rowType = "DETAIL" Then
            total = Application.WorksheetFunction.Sum( _
                ws.Range(ws.Cells(r, COL_FIRST_DAY), ws.Cells(r, COL_FIRST_DAY + daysInMonth - 1)))
            WriteBlankIfZero ws.Cells(r, COL_MONTH_TOTAL), total

            If daysInMonth < 31 Then
                ws.Range(ws.Cells(r, COL_FIRST_DAY + daysInMonth), ws.Cells(r, COL_LAST_DAY)).ClearContents
            End If

        ElseIf rowType = "SUBTOTAL" Then
            ws.Range(ws.Cells(r, COL_FIRST_DAY), ws.Cells(r, COL_MONTH_TOTAL)).ClearContents
        End If
    Next r

    For Each subtotalGroup In subtotalRows.Keys
        Set subtotalList = subtotalRows(subtotalGroup)

        If groupDetails.Exists(CStr(subtotalGroup)) Then
            Set detailList = groupDetails(CStr(subtotalGroup))
        Else
            Set detailList = Nothing
        End If

        For Each subtotalRow In subtotalList
            total = 0

            For dayIndex = 1 To daysInMonth
                dayTotal = 0

                If Not detailList Is Nothing Then
                    For Each detailRow In detailList
                        dayTotal = dayTotal + CDbl(Val(ws.Cells(CLng(detailRow), COL_FIRST_DAY + dayIndex - 1).Value))
                    Next detailRow
                End If

                WriteBlankIfZero ws.Cells(CLng(subtotalRow), COL_FIRST_DAY + dayIndex - 1), dayTotal
                total = total + dayTotal
            Next dayIndex

            WriteBlankIfZero ws.Cells(CLng(subtotalRow), COL_MONTH_TOTAL), total

            If daysInMonth < 31 Then
                ws.Range(ws.Cells(CLng(subtotalRow), COL_FIRST_DAY + daysInMonth), _
                         ws.Cells(CLng(subtotalRow), COL_LAST_DAY)).ClearContents
            End If
        Next subtotalRow
    Next subtotalGroup
End Sub

Private Sub WriteBlankIfZero(ByVal targetCell As Range, ByVal valueToWrite As Double)
    If valueToWrite = 0 Then
        targetCell.ClearContents
    Else
        targetCell.Value = valueToWrite
    End If
End Sub

Private Function GetLastUsedRow(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long) As Long
    Dim c As Long
    Dim candidate As Long

    GetLastUsedRow = ROW_HEADER

    For c = firstCol To lastCol
        candidate = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If candidate > GetLastUsedRow Then GetLastUsedRow = candidate
    Next c
End Function

Private Function LastRowInColumn(ByVal ws As Worksheet, ByVal colNumber As Long) As Long
    LastRowInColumn = ws.Cells(ws.Rows.Count, colNumber).End(xlUp).Row
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a >= b Then
        MaxLong = a
    Else
        MaxLong = b
    End If
End Function

Private Function NormalizeCode(ByVal value As Variant) As String
    If IsError(value) Then
        NormalizeCode = ""
    Else
        NormalizeCode = Trim$(CStr(value))
    End If
End Function

Private Function ValueForLog(ByVal value As Variant) As String
    If IsError(value) Then
        ValueForLog = "#ERROR"
    Else
        ValueForLog = Trim$(CStr(value))
    End If
End Function

Private Function MakeKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeKey = branchCode & KEY_SEP & businessCode
End Function

```
