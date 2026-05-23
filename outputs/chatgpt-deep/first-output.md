# chatgpt-deep 初回出力

```vba
Option Explicit

'============================================================
' 月次クレーム集計マクロ
'
' 入口:
'   RunMonthlyClaimAggregation
'
' 前提:
'   ThisWorkbook.Name = 月次クレーム集計YYMM.xlsm
'   日次フォルダ      = ThisWorkbook.Path & "\daily\YYMM\"
'   日次ファイル名    = クレーム集計YYMMDD.xlsx
'
' 追加参照:
'   不要
'   Scripting.Dictionary / FileSystemObject は late binding で使用
'============================================================

Private Const SHEET_MONTHLY As String = "月次集計"
Private Const SHEET_MASTER As String = "マスタ"
Private Const SHEET_LOG As String = "異常系"
Private Const SHEET_DAILY As String = "日次集計"

Private Const MONTHLY_PREFIX As String = "月次クレーム集計"
Private Const MONTHLY_SUFFIX As String = ".xlsm"

Private Const DAILY_PREFIX As String = "クレーム集計"
Private Const DAILY_SUFFIX As String = ".xlsx"

Private Const ROW_TYPE_DETAIL As String = "DETAIL"
Private Const ROW_TYPE_SUBTOTAL As String = "SUBTOTAL"

Private Const COL_ROW_TYPE As Long = 1          ' A
Private Const COL_SUBTOTAL_GROUP As Long = 2    ' B
Private Const COL_REGION As Long = 3            ' C
Private Const COL_BRANCH_CODE As Long = 4       ' D
Private Const COL_BRANCH_NAME As Long = 5       ' E
Private Const COL_BUSINESS_CODE As Long = 6     ' F
Private Const COL_BUSINESS_NAME As Long = 7     ' G
Private Const COL_FIRST_DAY As Long = 8         ' H
Private Const COL_LAST_DAY As Long = 38         ' AL
Private Const COL_MONTH_TOTAL As Long = 39      ' AM
Private Const COL_NOTE As Long = 40             ' AN

Private Const DAILY_COL_PROCESS_TYPE As Long = 1
Private Const DAILY_COL_BRANCH_CODE As Long = 2
Private Const DAILY_COL_BRANCH_NAME As Long = 3
Private Const DAILY_COL_BUSINESS_CODE As Long = 4
Private Const DAILY_COL_BUSINESS_NAME As Long = 5
Private Const DAILY_COL_CLAIM_COUNT As Long = 6
Private Const DAILY_COL_NOTE As Long = 7

Private Const KEY_DELIMITER As String = vbTab

Private gLogRow As Long
Private gIssueCount As Long
Private gProcessedDailyFileCount As Long
Private gTransferredRowCount As Long

'============================================================
' 入口プロシージャ
'============================================================
Public Sub RunMonthlyClaimAggregation()

    Dim wbMonthly As Workbook
    Dim wsMonthly As Worksheet
    Dim wsMaster As Worksheet
    Dim wsLog As Worksheet

    Dim targetYYMM As String
    Dim targetYear As Long
    Dim targetMonth As Long
    Dim daysInMonth As Long
    Dim targetFirstDate As Date

    Dim basePath As String
    Dim dailyRootPath As String

    Dim detailRows As Object
    Dim knownBranches As Object
    Dim knownBusinesses As Object
    Dim dailyFilesByDate As Object

    Dim lastMonthlyRow As Long
    Dim message As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim appStateChanged As Boolean

    On Error GoTo FatalError

    Set wbMonthly = ThisWorkbook

    Set wsLog = GetWorksheetOrNothing(wbMonthly, SHEET_LOG)
    If wsLog Is Nothing Then
        MsgBox "シート '" & SHEET_LOG & "' がありません。処理を中止します。", vbCritical
        Exit Sub
    End If

    InitializeLogSheet wsLog

    If Not TryParseMonthlyWorkbookName(wbMonthly.Name, targetYYMM, targetYear, targetMonth, daysInMonth, message) Then
        LogIssue wsLog, "ERROR", "INVALID_MONTHLY_WORKBOOK_NAME", "", "", wbMonthly.Name, 0, "", "", "", "", message, ""
        MsgBox "月次ブック名が想定形式ではありません。" & vbCrLf & message, vbCritical
        Exit Sub
    End If

    If Len(wbMonthly.Path) = 0 Then
        LogIssue wsLog, "ERROR", "MONTHLY_WORKBOOK_NOT_SAVED", targetYYMM, "", wbMonthly.Name, 0, "", "", "", "", _
                 "ThisWorkbook.Path を取得できません。月次ブックを保存してから実行してください。", ""
        MsgBox "月次ブックが未保存です。保存してから実行してください。", vbCritical
        Exit Sub
    End If

    Set wsMonthly = GetWorksheetOrNothing(wbMonthly, SHEET_MONTHLY)
    Set wsMaster = GetWorksheetOrNothing(wbMonthly, SHEET_MASTER)

    If wsMonthly Is Nothing Then
        LogIssue wsLog, "ERROR", "MONTHLY_SHEET_NOT_FOUND", targetYYMM, "", wbMonthly.Name, 0, "", "", "", "", _
                 "シート '" & SHEET_MONTHLY & "' がありません。", wbMonthly.Path
        MsgBox "シート '" & SHEET_MONTHLY & "' がありません。処理を中止します。", vbCritical
        Exit Sub
    End If

    If wsMaster Is Nothing Then
        LogIssue wsLog, "ERROR", "MASTER_SHEET_NOT_FOUND", targetYYMM, "", wbMonthly.Name, 0, "", "", "", "", _
                 "シート '" & SHEET_MASTER & "' がありません。", wbMonthly.Path
        MsgBox "シート '" & SHEET_MASTER & "' がありません。処理を中止します。", vbCritical
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    appStateChanged = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    basePath = wbMonthly.Path
    dailyRootPath = basePath & "\daily\" & targetYYMM
    targetFirstDate = DateSerial(targetYear, targetMonth, 1)

    Set detailRows = CreateDictionary()
    Set knownBranches = CreateDictionary()
    Set knownBusinesses = CreateDictionary()
    Set dailyFilesByDate = CreateDictionary()

    lastMonthlyRow = LastUsedRowInColumns(wsMonthly, COL_ROW_TYPE, COL_BUSINESS_NAME)

    BuildMonthlyDetailIndex wsMonthly, wsLog, targetYYMM, basePath, detailRows, lastMonthlyRow
    BuildMasterCodeSets wsMaster, wsMonthly, wsLog, targetYYMM, basePath, knownBranches, knownBusinesses
    ClearMonthlyTransferArea wsMonthly, lastMonthlyRow

    If Not FolderExists(dailyRootPath) Then
        LogIssue wsLog, "ERROR", "DAILY_FOLDER_NOT_FOUND", targetYYMM, "", dailyRootPath, 0, "", "", "", "", _
                 "日次ファイル探索起点フォルダが存在しません。", basePath
    Else
        EnumerateDailyFiles dailyRootPath, targetYYMM, basePath, wsLog, dailyFilesByDate
        ProcessDailyFileGroups dailyFilesByDate, wsMonthly, wsLog, targetYYMM, basePath, detailRows, knownBranches, knownBusinesses
    End If

    RecalculateMonthlyTotalsAndSubtotals wsMonthly, wsLog, targetYYMM, basePath, lastMonthlyRow, daysInMonth

    MsgBox "月次クレーム集計が完了しました。" & vbCrLf & _
           "対象月: " & targetYYMM & vbCrLf & _
           "処理日次ファイル数: " & gProcessedDailyFileCount & vbCrLf & _
           "転記行数: " & gTransferredRowCount & vbCrLf & _
           "ログ件数: " & gIssueCount, vbInformation

CleanExit:
    If appStateChanged Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.DisplayAlerts = oldDisplayAlerts
        Application.Calculation = oldCalculation
    End If
    Exit Sub

FatalError:
    On Error Resume Next
    LogIssue wsLog, "ERROR", "UNHANDLED_ERROR", targetYYMM, "", wbMonthly.Name, 0, "", "", "", "", _
             "予期しないエラー: " & Err.Number & " / " & Err.Description, basePath
    On Error GoTo 0

    MsgBox "予期しないエラーで処理を中止しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical

    Resume CleanExit

End Sub

'============================================================
' 月次表の初期化・索引作成
'============================================================
Private Sub BuildMonthlyDetailIndex( _
    ByVal wsMonthly As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal detailRows As Object, _
    ByRef lastMonthlyRow As Long)

    Dim r As Long
    Dim rowType As String
    Dim branchCode As String
    Dim businessCode As String
    Dim key As String

    If lastMonthlyRow < 2 Then
        LogIssue wsLog, "ERROR", "MONTHLY_DETAIL_NOT_FOUND", targetYYMM, "", ThisWorkbook.Name, 0, "", "", "", "", _
                 "月次集計シートに DETAIL 行が見つかりません。", basePath
        Exit Sub
    End If

    For r = 2 To lastMonthlyRow
        rowType = UCase$(NormalizeTextFromCell(wsMonthly.Cells(r, COL_ROW_TYPE)))

        If rowType = ROW_TYPE_DETAIL Then
            branchCode = NormalizeCodeFromCell(wsMonthly.Cells(r, COL_BRANCH_CODE))
            businessCode = NormalizeCodeFromCell(wsMonthly.Cells(r, COL_BUSINESS_CODE))

            If Len(branchCode) = 0 Then
                LogIssue wsLog, "ERROR", "MONTHLY_DETAIL_BRANCH_CODE_BLANK", targetYYMM, "", ThisWorkbook.Name, r, "", businessCode, "", "", _
                         "月次集計 DETAIL 行の支店コードが空欄です。", basePath
            End If

            If Len(businessCode) = 0 Then
                LogIssue wsLog, "ERROR", "MONTHLY_DETAIL_BUSINESS_CODE_BLANK", targetYYMM, "", ThisWorkbook.Name, r, branchCode, "", "", "", _
                         "月次集計 DETAIL 行の業務コードが空欄です。", basePath
            End If

            If Len(branchCode) > 0 And Len(businessCode) > 0 Then
                key = MakeTransferKey(branchCode, businessCode)

                If detailRows.Exists(key) Then
                    LogIssue wsLog, "ERROR", "DUPLICATE_MONTHLY_DETAIL_KEY", targetYYMM, "", ThisWorkbook.Name, r, branchCode, businessCode, "", "", _
                             "月次集計に同じ支店コード + 業務コードの DETAIL 行が複数あります。先頭行を転記先として使用します。", basePath
                Else
                    detailRows.Add key, r
                End If
            End If
        End If
    Next r

    If detailRows.Count = 0 Then
        LogIssue wsLog, "ERROR", "MONTHLY_DETAIL_KEY_NOT_FOUND", targetYYMM, "", ThisWorkbook.Name, 0, "", "", "", "", _
                 "有効な DETAIL 転記キーが作成できませんでした。", basePath
    End If

End Sub

Private Sub ClearMonthlyTransferArea(ByVal wsMonthly As Worksheet, ByVal lastMonthlyRow As Long)

    Dim r As Long
    Dim rowType As String

    For r = 2 To lastMonthlyRow
        rowType = UCase$(NormalizeTextFromCell(wsMonthly.Cells(r, COL_ROW_TYPE)))

        If rowType = ROW_TYPE_DETAIL Or rowType = ROW_TYPE_SUBTOTAL Then
            wsMonthly.Range(wsMonthly.Cells(r, COL_FIRST_DAY), wsMonthly.Cells(r, COL_LAST_DAY)).ClearContents
            wsMonthly.Cells(r, COL_MONTH_TOTAL).ClearContents
        End If
    Next r

End Sub

Private Sub RecalculateMonthlyTotalsAndSubtotals( _
    ByVal wsMonthly As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal lastMonthlyRow As Long, _
    ByVal daysInMonth As Long)

    Dim r As Long
    Dim d As Long
    Dim c As Long
    Dim detailRow As Long

    Dim rowType As String
    Dim groupName As String
    Dim detailGroupName As String

    Dim dayTotal As Double
    Dim monthTotal As Double
    Dim foundDetail As Boolean
    Dim v As Variant

    ' DETAIL 行の AM 月合計
    For r = 2 To lastMonthlyRow
        rowType = UCase$(NormalizeTextFromCell(wsMonthly.Cells(r, COL_ROW_TYPE)))

        If rowType = ROW_TYPE_DETAIL Then
            monthTotal = 0

            For d = 1 To daysInMonth
                c = COL_FIRST_DAY + d - 1
                v = wsMonthly.Cells(r, c).Value2
                If IsNumeric(v) Then
                    monthTotal = monthTotal + CDbl(v)
                End If
            Next d

            wsMonthly.Cells(r, COL_MONTH_TOTAL).Value = monthTotal
        End If
    Next r

    ' SUBTOTAL 行の日別・月合計
    For r = 2 To lastMonthlyRow
        rowType = UCase$(NormalizeTextFromCell(wsMonthly.Cells(r, COL_ROW_TYPE)))

        If rowType = ROW_TYPE_SUBTOTAL Then
            groupName = NormalizeTextFromCell(wsMonthly.Cells(r, COL_SUBTOTAL_GROUP))

            If Len(groupName) = 0 Then
                LogIssue wsLog, "ERROR", "SUBTOTAL_GROUP_BLANK", targetYYMM, "", ThisWorkbook.Name, r, "", "", "", "", _
                         "SUBTOTAL 行の小計グループが空欄です。小計計算をスキップします。", basePath
            Else
                monthTotal = 0
                foundDetail = False

                For d = 1 To daysInMonth
                    c = COL_FIRST_DAY + d - 1
                    dayTotal = 0

                    For detailRow = 2 To lastMonthlyRow
                        If UCase$(NormalizeTextFromCell(wsMonthly.Cells(detailRow, COL_ROW_TYPE))) = ROW_TYPE_DETAIL Then
                            detailGroupName = NormalizeTextFromCell(wsMonthly.Cells(detailRow, COL_SUBTOTAL_GROUP))

                            If StrComp(detailGroupName, groupName, vbTextCompare) = 0 Then
                                foundDetail = True
                                v = wsMonthly.Cells(detailRow, c).Value2
                                If IsNumeric(v) Then
                                    dayTotal = dayTotal + CDbl(v)
                                End If
                            End If
                        End If
                    Next detailRow

                    wsMonthly.Cells(r, c).Value = dayTotal
                    monthTotal = monthTotal + dayTotal
                Next d

                wsMonthly.Cells(r, COL_MONTH_TOTAL).Value = monthTotal

                If Not foundDetail Then
                    LogIssue wsLog, "WARNING", "SUBTOTAL_GROUP_WITHOUT_DETAIL", targetYYMM, "", ThisWorkbook.Name, r, "", "", "", "", _
                             "同じ小計グループの DETAIL 行がありません。小計は 0 になります。", basePath
                End If
            End If
        End If
    Next r

End Sub

'============================================================
' マスタ読込
'============================================================
Private Sub BuildMasterCodeSets( _
    ByVal wsMaster As Worksheet, _
    ByVal wsMonthly As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal knownBranches As Object, _
    ByVal knownBusinesses As Object)

    Dim branchHeaderRow As Long
    Dim businessHeaderRow As Long
    Dim branchCol As Long
    Dim businessCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim code As String

    branchCol = FindHeaderColumn(wsMaster, "支店コード", branchHeaderRow)
    businessCol = FindHeaderColumn(wsMaster, "業務コード", businessHeaderRow)

    If branchCol > 0 Then
        lastRow = LastUsedRowInColumns(wsMaster, branchCol, branchCol)
        For r = branchHeaderRow + 1 To lastRow
            code = NormalizeCodeFromCell(wsMaster.Cells(r, branchCol))
            If Len(code) > 0 Then
                If Not knownBranches.Exists(code) Then knownBranches.Add code, True
            End If
        Next r
    Else
        LogIssue wsLog, "WARNING", "MASTER_BRANCH_HEADER_NOT_FOUND", targetYYMM, "", ThisWorkbook.Name, 0, "", "", "", "", _
                 "マスタシートから '支店コード' 見出しを検出できません。月次集計 DETAIL 行から支店コード集合を補完します。", basePath
    End If

    If businessCol > 0 Then
        lastRow = LastUsedRowInColumns(wsMaster, businessCol, businessCol)
        For r = businessHeaderRow + 1 To lastRow
            code = NormalizeCodeFromCell(wsMaster.Cells(r, businessCol))
            If Len(code) > 0 Then
                If Not knownBusinesses.Exists(code) Then knownBusinesses.Add code, True
            End If
        Next r
    Else
        LogIssue wsLog, "WARNING", "MASTER_BUSINESS_HEADER_NOT_FOUND", targetYYMM, "", ThisWorkbook.Name, 0, "", "", "", "", _
                 "マスタシートから '業務コード' 見出しを検出できません。月次集計 DETAIL 行から業務コード集合を補完します。", basePath
    End If

    If knownBranches.Count = 0 Then
        AddMonthlyDetailCodesToSet wsMonthly, COL_BRANCH_CODE, knownBranches
        LogIssue wsLog, "WARNING", "MASTER_BRANCH_FALLBACK_USED", targetYYMM, "", ThisWorkbook.Name, 0, "", "", "", "", _
                 "有効な支店マスタを取得できなかったため、月次集計 DETAIL 行の支店コードを既知支店として扱います。", basePath
    End If

    If knownBusinesses.Count = 0 Then
        AddMonthlyDetailCodesToSet wsMonthly, COL_BUSINESS_CODE, knownBusinesses
        LogIssue wsLog, "WARNING", "MASTER_BUSINESS_FALLBACK_USED", targetYYMM, "", ThisWorkbook.Name, 0, "", "", "", "", _
                 "有効な業務マスタを取得できなかったため、月次集計 DETAIL 行の業務コードを既知業務として扱います。", basePath
    End If

End Sub

Private Sub AddMonthlyDetailCodesToSet(ByVal wsMonthly As Worksheet, ByVal targetCol As Long, ByVal codeSet As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim rowType As String
    Dim code As String

    lastRow = LastUsedRowInColumns(wsMonthly, COL_ROW_TYPE, COL_BUSINESS_NAME)

    For r = 2 To lastRow
        rowType = UCase$(NormalizeTextFromCell(wsMonthly.Cells(r, COL_ROW_TYPE)))
        If rowType = ROW_TYPE_DETAIL Then
            code = NormalizeCodeFromCell(wsMonthly.Cells(r, targetCol))
            If Len(code) > 0 Then
                If Not codeSet.Exists(code) Then codeSet.Add code, True
            End If
        End If
    Next r

End Sub

'============================================================
' 日次ファイル探索・処理
'============================================================
Private Sub EnumerateDailyFiles( _
    ByVal rootPath As String, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal wsLog As Worksheet, _
    ByVal dailyFilesByDate As Object)

    Dim fso As Object
    Dim rootFolder As Object

    On Error GoTo EnumerateError

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(rootPath)

    EnumerateDailyFolder rootFolder, targetYYMM, basePath, wsLog, dailyFilesByDate

    Exit Sub

EnumerateError:
    LogIssue wsLog, "ERROR", "DAILY_FOLDER_ENUMERATION_FAILED", targetYYMM, "", rootPath, 0, "", "", "", "", _
             "日次フォルダの再帰探索に失敗しました: " & Err.Number & " / " & Err.Description, basePath

End Sub

Private Sub EnumerateDailyFolder( _
    ByVal folder As Object, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal wsLog As Worksheet, _
    ByVal dailyFilesByDate As Object)

    Dim fileItem As Object
    Dim subFolder As Object
    Dim ext As String

    Dim dateKey As String
    Dim fileYYMM As String
    Dim fileDate As Date
    Dim message As String
    Dim files As Collection

    On Error GoTo FolderError

    For Each fileItem In folder.Files
        If Left$(fileItem.Name, 2) <> "~$" Then
            ext = LCase$(Mid$(fileItem.Name, InStrRev(fileItem.Name, ".")))

            If Left$(ext, 4) = ".xls" Then
                If TryParseDailyFileName(fileItem.Name, dateKey, fileDate, fileYYMM, message) Then
                    If StrComp(fileYYMM, targetYYMM, vbTextCompare) <> 0 Then
                        LogIssue wsLog, "WARNING", "OUT_OF_MONTH_DAILY_FILE", targetYYMM, fileDate, fileItem.Path, 0, "", "", "", "", _
                                 "対象月と異なる日次ファイルです。処理対象外にしました。", basePath
                    Else
                        If Not dailyFilesByDate.Exists(dateKey) Then
                            Set files = New Collection
                            dailyFilesByDate.Add dateKey, files
                        End If

                        Set files = dailyFilesByDate.Item(dateKey)
                        files.Add CStr(fileItem.Path)
                    End If
                Else
                    LogIssue wsLog, "WARNING", "INVALID_DAILY_FILE_NAME", targetYYMM, "", fileItem.Path, 0, "", "", "", "", _
                             message, basePath
                End If
            End If
        End If
    Next fileItem

    For Each subFolder In folder.SubFolders
        EnumerateDailyFolder subFolder, targetYYMM, basePath, wsLog, dailyFilesByDate
    Next subFolder

    Exit Sub

FolderError:
    LogIssue wsLog, "ERROR", "DAILY_FOLDER_ENUMERATION_FAILED", targetYYMM, "", folder.Path, 0, "", "", "", "", _
             "フォルダ探索中にエラーが発生しました: " & Err.Number & " / " & Err.Description, basePath

End Sub

Private Sub ProcessDailyFileGroups( _
    ByVal dailyFilesByDate As Object, _
    ByVal wsMonthly As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal detailRows As Object, _
    ByVal knownBranches As Object, _
    ByVal knownBusinesses As Object)

    Dim dateKeys As Variant
    Dim filePaths As Variant
    Dim dateKey As String
    Dim fileDate As Date
    Dim dayNumber As Long
    Dim i As Long
    Dim j As Long
    Dim files As Collection

    If dailyFilesByDate.Count = 0 Then
        LogIssue wsLog, "WARNING", "NO_TARGET_DAILY_FILE", targetYYMM, "", ThisWorkbook.Path & "\daily\" & targetYYMM, 0, "", "", "", "", _
                 "対象月の日次ファイルが見つかりませんでした。", basePath
        Exit Sub
    End If

    dateKeys = dailyFilesByDate.Keys
    SortVariantStringArray dateKeys

    For i = LBound(dateKeys) To UBound(dateKeys)
        dateKey = CStr(dateKeys(i))
        Set files = dailyFilesByDate.Item(dateKey)

        filePaths = CollectionToVariantArray(files)
        SortVariantStringArray filePaths

        fileDate = DateSerial(2000 + CLng(Left$(dateKey, 2)), CLng(Mid$(dateKey, 3, 2)), CLng(Right$(dateKey, 2)))
        dayNumber = Day(fileDate)

        ' 同じ日付のファイルが複数ある場合:
        ' フルパス昇順の先頭1件だけを処理し、残りは重複としてログ
        ProcessDailyWorkbook CStr(filePaths(LBound(filePaths))), fileDate, dayNumber, wsMonthly, wsLog, targetYYMM, basePath, detailRows, knownBranches, knownBusinesses

        If UBound(filePaths) > LBound(filePaths) Then
            For j = LBound(filePaths) + 1 To UBound(filePaths)
                LogIssue wsLog, "WARNING", "DUPLICATE_DAILY_FILE", targetYYMM, fileDate, CStr(filePaths(j)), 0, "", "", "", "", _
                         "同じ日付の日次ファイルが複数あります。フルパス昇順の先頭1件のみ処理し、このファイルは処理対象外にしました。", basePath
            Next j
        End If
    Next i

End Sub

Private Sub ProcessDailyWorkbook( _
    ByVal filePath As String, _
    ByVal fileDate As Date, _
    ByVal dayNumber As Long, _
    ByVal wsMonthly As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal detailRows As Object, _
    ByVal knownBranches As Object, _
    ByVal knownBusinesses As Object)

    Dim wbDaily As Workbook
    Dim wsDaily As Worksheet
    Dim lastRow As Long
    Dim r As Long

    On Error GoTo DailyFileError

    Set wbDaily = Application.Workbooks.Open( _
        Filename:=filePath, _
        UpdateLinks:=False, _
        ReadOnly:=True, _
        AddToMru:=False)

    Set wsDaily = GetWorksheetOrNothing(wbDaily, SHEET_DAILY)

    If wsDaily Is Nothing Then
        LogIssue wsLog, "ERROR", "DAILY_SHEET_NOT_FOUND", targetYYMM, fileDate, filePath, 0, "", "", "", "", _
                 "日次ブックにシート '" & SHEET_DAILY & "' がありません。", basePath
        GoTo CleanUp
    End If

    gProcessedDailyFileCount = gProcessedDailyFileCount + 1

    lastRow = LastUsedRowInColumns(wsDaily, DAILY_COL_PROCESS_TYPE, DAILY_COL_NOTE)

    If lastRow >= 2 Then
        For r = 2 To lastRow
            ProcessDailyRow wsDaily, r, filePath, fileDate, dayNumber, wsMonthly, wsLog, targetYYMM, basePath, detailRows, knownBranches, knownBusinesses
        Next r
    End If

CleanUp:
    On Error Resume Next
    If Not wbDaily Is Nothing Then wbDaily.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

DailyFileError:
    LogIssue wsLog, "ERROR", "DAILY_WORKBOOK_PROCESS_FAILED", targetYYMM, fileDate, filePath, 0, "", "", "", "", _
             "日次ブックのオープンまたは処理に失敗しました: " & Err.Number & " / " & Err.Description, basePath
    Resume CleanUp

End Sub

Private Sub ProcessDailyRow( _
    ByVal wsDaily As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal filePath As String, _
    ByVal fileDate As Date, _
    ByVal dayNumber As Long, _
    ByVal wsMonthly As Worksheet, _
    ByVal wsLog As Worksheet, _
    ByVal targetYYMM As String, _
    ByVal basePath As String, _
    ByVal detailRows As Object, _
    ByVal knownBranches As Object, _
    ByVal knownBusinesses As Object)

    Dim processType As String
    Dim branchCode As String
    Dim businessCode As String
    Dim claimCountText As String
    Dim claimCount As Double
    Dim countIssueType As String
    Dim countIssueMessage As String

    Dim isOkRow As Boolean
    Dim hasIssue As Boolean
    Dim key As String
    Dim targetRow As Long

    processType = NormalizeTextFromCell(wsDaily.Cells(rowNumber, DAILY_COL_PROCESS_TYPE))
    branchCode = NormalizeCodeFromCell(wsDaily.Cells(rowNumber, DAILY_COL_BRANCH_CODE))
    businessCode = NormalizeCodeFromCell(wsDaily.Cells(rowNumber, DAILY_COL_BUSINESS_CODE))

    isOkRow = (StrComp(processType, "ok", vbTextCompare) = 0)

    If Len(branchCode) = 0 Then
        hasIssue = True
        LogIssue wsLog, "ERROR", "BLANK_BRANCH_CODE", targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, "", processType, _
                 "支店コードが空欄です。転記対象外にしました。", basePath
    ElseIf Not knownBranches.Exists(branchCode) Then
        hasIssue = True
        LogIssue wsLog, "ERROR", "UNKNOWN_BRANCH_CODE", targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, "", processType, _
                 "マスタに存在しない支店コードです。転記対象外にしました。", basePath
    End If

    If Len(businessCode) = 0 Then
        hasIssue = True
        LogIssue wsLog, "ERROR", "BLANK_BUSINESS_CODE", targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, "", processType, _
                 "業務コードが空欄です。転記対象外にしました。", basePath
    ElseIf Not knownBusinesses.Exists(businessCode) Then
        hasIssue = True
        LogIssue wsLog, "ERROR", "UNKNOWN_BUSINESS_CODE", targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, "", processType, _
                 "マスタに存在しない業務コードです。転記対象外にしました。", basePath
    End If

    If Len(branchCode) > 0 And Len(businessCode) > 0 Then
        key = MakeTransferKey(branchCode, businessCode)

        If Not detailRows.Exists(key) Then
            hasIssue = True
            LogIssue wsLog, "ERROR", "MONTHLY_DETAIL_KEY_NOT_FOUND", targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, "", processType, _
                     "月次集計に該当する支店コード + 業務コードの DETAIL 行がありません。転記対象外にしました。", basePath
        End If
    End If

    If Not TryReadClaimCount(wsDaily.Cells(rowNumber, DAILY_COL_CLAIM_COUNT), claimCount, claimCountText, countIssueType, countIssueMessage) Then
        hasIssue = True
        LogIssue wsLog, "ERROR", countIssueType, targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, claimCountText, processType, _
                 countIssueMessage, basePath
    End If

    If Not isOkRow Then
        If Not hasIssue Then
            LogIssue wsLog, "WARNING", "NON_OK_ROW", targetYYMM, fileDate, filePath, rowNumber, branchCode, businessCode, claimCountText, processType, _
                     "処理区分が 'ok' ではないため、異常系サンプルとして転記対象外にしました。", basePath
        End If
        Exit Sub
    End If

    If hasIssue Then
        Exit Sub
    End If

    targetRow = CLng(detailRows.Item(key))
    AddClaimCountToMonthly wsMonthly, targetRow, dayNumber, claimCount
    gTransferredRowCount = gTransferredRowCount + 1

End Sub

Private Sub AddClaimCountToMonthly( _
    ByVal wsMonthly As Worksheet, _
    ByVal targetRow As Long, _
    ByVal dayNumber As Long, _
    ByVal claimCount As Double)

    Dim targetCol As Long
    Dim currentValue As Variant

    targetCol = COL_FIRST_DAY + dayNumber - 1
    currentValue = wsMonthly.Cells(targetRow, targetCol).Value2

    If IsNumeric(currentValue) Then
        wsMonthly.Cells(targetRow, targetCol).Value = CDbl(currentValue) + claimCount
    Else
        wsMonthly.Cells(targetRow, targetCol).Value = claimCount
    End If

End Sub

'============================================================
' ファイル名解析
'============================================================
Private Function TryParseMonthlyWorkbookName( _
    ByVal workbookName As String, _
    ByRef targetYYMM As String, _
    ByRef targetYear As Long, _
    ByRef targetMonth As Long, _
    ByRef daysInMonth As Long, _
    ByRef message As String) As Boolean

    Dim body As String
    Dim yy As Long
    Dim mm As Long

    If Left$(workbookName, Len(MONTHLY_PREFIX)) <> MONTHLY_PREFIX Then
        message = "月次ブック名の接頭辞が '" & MONTHLY_PREFIX & "' ではありません。"
        Exit Function
    End If

    If LCase$(Right$(workbookName, Len(MONTHLY_SUFFIX))) <> MONTHLY_SUFFIX Then
        message = "月次ブックの拡張子が '" & MONTHLY_SUFFIX & "' ではありません。"
        Exit Function
    End If

    body = Mid$(workbookName, Len(MONTHLY_PREFIX) + 1, Len(workbookName) - Len(MONTHLY_PREFIX) - Len(MONTHLY_SUFFIX))

    If Len(body) <> 4 Or Not IsAllDigits(body) Then
        message = "月次ブック名の YYMM 部分が4桁数字ではありません。"
        Exit Function
    End If

    yy = CLng(Left$(body, 2))
    mm = CLng(Right$(body, 2))

    If mm < 1 Or mm > 12 Then
        message = "月次ブック名の MM が 01～12 の範囲外です。"
        Exit Function
    End If

    targetYYMM = body
    targetYear = 2000 + yy
    targetMonth = mm
    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))

    TryParseMonthlyWorkbookName = True

End Function

Private Function TryParseDailyFileName( _
    ByVal fileName As String, _
    ByRef dateKey As String, _
    ByRef fileDate As Date, _
    ByRef fileYYMM As String, _
    ByRef message As String) As Boolean

    Dim body As String
    Dim yy As Long
    Dim mm As Long
    Dim dd As Long
    Dim yyyy As Long

    If Left$(fileName, Len(DAILY_PREFIX)) <> DAILY_PREFIX Then
        message = "日次ファイル名の接頭辞が '" & DAILY_PREFIX & "' ではありません。"
        Exit Function
    End If

    If LCase$(Right$(fileName, Len(DAILY_SUFFIX))) <> DAILY_SUFFIX Then
        message = "日次ファイルの拡張子が '" & DAILY_SUFFIX & "' ではありません。"
        Exit Function
    End If

    body = Mid$(fileName, Len(DAILY_PREFIX) + 1, Len(fileName) - Len(DAILY_PREFIX) - Len(DAILY_SUFFIX))

    If Len(body) <> 6 Or Not IsAllDigits(body) Then
        message = "日次ファイル名の YYMMDD 部分が6桁数字ではありません。"
        Exit Function
    End If

    yy = CLng(Left$(body, 2))
    mm = CLng(Mid$(body, 3, 2))
    dd = CLng(Right$(body, 2))
    yyyy = 2000 + yy

    If Not TryMakeDate(yyyy, mm, dd, fileDate) Then
        message = "日次ファイル名の YYMMDD が有効な日付ではありません。"
        Exit Function
    End If

    dateKey = body
    fileYYMM = Left$(body, 4)

    TryParseDailyFileName = True

End Function

Private Function TryMakeDate(ByVal yyyy As Long, ByVal mm As Long, ByVal dd As Long, ByRef resultDate As Date) As Boolean

    On Error GoTo InvalidDate

    resultDate = DateSerial(yyyy, mm, dd)

    If Year(resultDate) <> yyyy Then Exit Function
    If Month(resultDate) <> mm Then Exit Function
    If Day(resultDate) <> dd Then Exit Function

    TryMakeDate = True
    Exit Function

InvalidDate:
    TryMakeDate = False

End Function

'============================================================
' ログ
'============================================================
Private Sub InitializeLogSheet(ByVal wsLog As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "記録日時", _
        "重要度", _
        "異常コード", _
        "対象月", _
        "日付", _
        "ファイル", _
        "行番号", _
        "処理区分", _
        "支店コード", _
        "業務コード", _
        "クレーム件数", _
        "内容")

    wsLog.Cells.ClearContents

    For i = LBound(headers) To UBound(headers)
        wsLog.Cells(1, i + 1).Value = headers(i)
    Next i

    wsLog.Rows(1).Font.Bold = True

    gLogRow = 2
    gIssueCount = 0
    gProcessedDailyFileCount = 0
    gTransferredRowCount = 0

End Sub

Private Sub LogIssue( _
    ByVal wsLog As Worksheet, _
    ByVal severity As String, _
    ByVal issueType As String, _
    ByVal targetYYMM As String, _
    ByVal dailyDate As Variant, _
    ByVal filePath As String, _
    ByVal sheetRow As Long, _
    ByVal branchCode As String, _
    ByVal businessCode As String, _
    ByVal claimCountText As String, _
    ByVal processType As String, _
    ByVal message As String, _
    ByVal basePath As String)

    If wsLog Is Nothing Then Exit Sub

    With wsLog
        .Cells(gLogRow, 1).Value = Now
        .Cells(gLogRow, 2).Value = severity
        .Cells(gLogRow, 3).Value = issueType
        .Cells(gLogRow, 4).Value = targetYYMM

        If IsDate(dailyDate) Then
            .Cells(gLogRow, 5).Value = CDate(dailyDate)
            .Cells(gLogRow, 5).NumberFormatLocal = "yyyy/m/d"
        Else
            .Cells(gLogRow, 5).Value = ""
        End If

        .Cells(gLogRow, 6).Value = GetRelativePath(filePath, basePath)

        If sheetRow > 0 Then
            .Cells(gLogRow, 7).Value = sheetRow
        Else
            .Cells(gLogRow, 7).Value = ""
        End If

        .Cells(gLogRow, 8).Value = processType
        .Cells(gLogRow, 9).Value = branchCode
        .Cells(gLogRow, 10).Value = businessCode
        .Cells(gLogRow, 11).Value = claimCountText
        .Cells(gLogRow, 12).Value = message
    End With

    gLogRow = gLogRow + 1
    gIssueCount = gIssueCount + 1

End Sub

'============================================================
' 値の読取・正規化
'============================================================
Private Function TryReadClaimCount( _
    ByVal targetCell As Range, _
    ByRef claimCount As Double, _
    ByRef claimCountText As String, _
    ByRef issueType As String, _
    ByRef issueMessage As String) As Boolean

    Dim v As Variant
    Dim s As String

    v = targetCell.Value2

    If IsError(v) Then
        claimCountText = "#ERROR"
        issueType = "NON_NUMERIC_CLAIM_COUNT"
        issueMessage = "クレーム件数セルがエラー値です。転記対象外にしました。"
        Exit Function
    End If

    claimCountText = NormalizeTextFromValue(v)

    If Len(claimCountText) = 0 Then
        issueType = "BLANK_CLAIM_COUNT"
        issueMessage = "クレーム件数が空欄です。転記対象外にしました。"
        Exit Function
    End If

    s = claimCountText
    s = Replace(s, ",", "")
    s = Replace(s, "，", "")

    If Not IsNumeric(s) Then
        issueType = "NON_NUMERIC_CLAIM_COUNT"
        issueMessage = "クレーム件数が数値ではありません。転記対象外にしました。"
        Exit Function
    End If

    claimCount = CDbl(s)
    TryReadClaimCount = True

End Function

Private Function NormalizeCodeFromCell(ByVal targetCell As Range) As String

    Dim v As Variant
    Dim textValue As String

    v = targetCell.Value2

    If IsError(v) Or IsEmpty(v) Then
        NormalizeCodeFromCell = ""
        Exit Function
    End If

    On Error Resume Next
    textValue = CStr(targetCell.Text)
    On Error GoTo 0

    textValue = NormalizeTextFromValue(textValue)

    If Len(textValue) > 0 And InStr(1, textValue, "#", vbBinaryCompare) = 0 Then
        NormalizeCodeFromCell = textValue
    Else
        NormalizeCodeFromCell = NormalizeTextFromValue(v)
    End If

End Function

Private Function NormalizeTextFromCell(ByVal targetCell As Range) As String

    Dim v As Variant
    v = targetCell.Value2

    If IsError(v) Or IsEmpty(v) Then
        NormalizeTextFromCell = ""
    Else
        NormalizeTextFromCell = NormalizeTextFromValue(v)
    End If

End Function

Private Function NormalizeTextFromValue(ByVal v As Variant) As String

    Dim s As String

    If IsError(v) Or IsEmpty(v) Then
        NormalizeTextFromValue = ""
        Exit Function
    End If

    s = CStr(v)
    s = Replace(s, ChrW(&H3000), " ")
    s = Trim$(s)

    NormalizeTextFromValue = s

End Function

Private Function MakeTransferKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeTransferKey = branchCode & KEY_DELIMITER & businessCode
End Function

'============================================================
' 汎用関数
'============================================================
Private Function GetWorksheetOrNothing(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetWorksheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0

End Function

Private Function CreateDictionary() As Object

    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    CreateDictionary.CompareMode = vbTextCompare

End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FolderExists = fso.FolderExists(folderPath)

End Function

Private Function IsAllDigits(ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String

    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    IsAllDigits = True

End Function

Private Function LastUsedRowInColumns(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long) As Long

    Dim c As Long
    Dim r As Long
    Dim maxRow As Long

    maxRow = 1

    For c = firstCol To lastCol
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > maxRow Then maxRow = r
    Next c

    LastUsedRowInColumns = maxRow

End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = lastCell.Column
    End If

End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String, ByRef headerRow As Long) As Long

    Dim maxRow As Long
    Dim maxCol As Long
    Dim r As Long
    Dim c As Long
    Dim cellText As String

    maxRow = LastUsedRowInColumns(ws, 1, LastUsedColumn(ws))
    If maxRow > 20 Then maxRow = 20

    maxCol = LastUsedColumn(ws)

    For r = 1 To maxRow
        For c = 1 To maxCol
            cellText = NormalizeTextFromCell(ws.Cells(r, c))
            If StrComp(cellText, headerText, vbTextCompare) = 0 Then
                headerRow = r
                FindHeaderColumn = c
                Exit Function
            End If
        Next c
    Next r

    headerRow = 0
    FindHeaderColumn = 0

End Function

Private Function GetRelativePath(ByVal fullPath As String, ByVal basePath As String) As String

    Dim normalizedFull As String
    Dim normalizedBase As String

    If Len(basePath) = 0 Then
        GetRelativePath = fullPath
        Exit Function
    End If

    normalizedFull = Replace(fullPath, "/", "\")
    normalizedBase = Replace(basePath, "/", "\")

    If Right$(normalizedBase, 1) <> "\" Then
        normalizedBase = normalizedBase & "\"
    End If

    If StrComp(Left$(normalizedFull, Len(normalizedBase)), normalizedBase, vbTextCompare) = 0 Then
        GetRelativePath = Mid$(fullPath, Len(normalizedBase) + 1)
    Else
        GetRelativePath = fullPath
    End If

End Function

'============================================================
' 配列ソート
'============================================================
Private Function CollectionToVariantArray(ByVal sourceCollection As Collection) As Variant

    Dim arr() As Variant
    Dim i As Long

    ReDim arr(0 To sourceCollection.Count - 1)

    For i = 1 To sourceCollection.Count
        arr(i - 1) = sourceCollection.Item(i)
    Next i

    CollectionToVariantArray = arr

End Function

Private Sub SortVariantStringArray(ByRef arr As Variant)

    If Not IsArray(arr) Then Exit Sub
    If UBound(arr) <= LBound(arr) Then Exit Sub

    QuickSortVariantStringArray arr, LBound(arr), UBound(arr)

End Sub

Private Sub QuickSortVariantStringArray(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)

    Dim low As Long
    Dim high As Long
    Dim pivot As String
    Dim temp As Variant

    low = first
    high = last
    pivot = CStr(arr((first + last) \ 2))

    Do While low <= high
        Do While StrComp(CStr(arr(low)), pivot, vbTextCompare) < 0
            low = low + 1
        Loop

        Do While StrComp(CStr(arr(high)), pivot, vbTextCompare) > 0
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

    If first < high Then QuickSortVariantStringArray arr, first, high
    If low < last Then QuickSortVariantStringArray arr, low, last

End Sub
```
