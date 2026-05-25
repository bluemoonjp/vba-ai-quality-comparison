# codex-plan-xhigh VBA生成成果物

## 短い計画

1. 許可された仕様、制約、条件、標準回答、サンプルCSV、期待結果を確認する。
2. 月次ブック名から対象年月 `YYMM` を取得し、月次ブックと同じフォルダ配下の `daily/<YYMM>/` から日次ファイルを読む。
3. 月次表の支店コード+業務コードを辞書化し、日次データを同日同キーで加算転記する。
4. 転記前に日別列と月合計列をクリアして0初期化し、転記後に月合計を再計算する。
5. MsgBoxは使わず、処理結果とエラーをログシートに残す。

## 実装方針

- 入口プロシージャは `ImportDailyClaimsToMonthly`。
- 標準モジュールへ貼り付けて使う完全なVBAコード。
- 追加参照は不要。`Scripting.Dictionary` は `CreateObject` によるlate bindingで使う。
- 日次ファイルは `ThisWorkbook.Path\daily\<YYMM>\クレーム集計<YYMMDD>.xlsx` を対象にする。
- 対象月は月次ブック名に含まれる末尾側の4桁数字から取得する。例: `月次クレーム集計2605.xlsm` なら `2605`。
- 月次表は `月次集計` シート、日次表は `日次集計` シートを使う。
- 月次表の列は A=支店コード、C=業務コード、E:AI=1日から31日、AJ=月合計として扱う。
- 日次表の列は A=支店コード、C=業務コード、E=クレーム件数として扱う。

## 配置場所

ExcelのVBEで対象の月次ブックに標準モジュールを追加し、以下のコードを貼り付けます。`.bas` として保存する場合は、このコードブロックの中身だけを保存します。

## 実行方法

1. 月次ブックをマクロ有効形式で保存し、月次ブックと同じフォルダ配下に `daily/<YYMM>/` を置きます。
2. `daily/<YYMM>/` に `クレーム集計<YYMMDD>.xlsx` 形式の日次ファイルを置きます。
3. Excelから `ImportDailyClaimsToMonthly` を実行します。

## 参照設定要否

追加参照は不要です。辞書はlate bindingで作成するため、VBEの参照設定変更は不要です。

## ログ出力

`取込ログ` シートを作成またはクリアし、日時、レベル、対象、内容を一覧で出力します。処理開始、対象月、入力フォルダ、ファイル単位の結果、スキップ行、終了サマリ、致命的エラーを記録します。

## エラー処理

MsgBoxで停止せず、ログシートへ記録して終了または該当ファイルをスキップします。日次ブックを開いた後にエラーが起きた場合は、保存せずに閉じます。Application設定は終了時に復元します。

## 再実行安全性

転記前に月次表の E:AJ を `ClearContents` し、日別列と月合計列を0で初期化してから日次データを加算します。そのため、同じ月次ブックで複数回実行しても前回結果へ二重加算されません。

## VBAコード

```vba
Option Explicit

Private Const MONTHLY_SHEET_NAME As String = "月次集計"
Private Const DAILY_SHEET_NAME As String = "日次集計"
Private Const LOG_SHEET_NAME As String = "取込ログ"

Private Const FIRST_DATA_ROW As Long = 2
Private Const COL_BRANCH_CODE As Long = 1
Private Const COL_BUSINESS_CODE As Long = 3
Private Const COL_CLAIM_COUNT As Long = 5
Private Const DAY_START_COL As Long = 5
Private Const DAY_END_COL As Long = 35
Private Const TOTAL_COL As Long = 36
Private Const KEY_DELIMITER As String = "|"

Public Sub ImportDailyClaimsToMonthly()
    Dim monthlyBook As Workbook
    Dim monthlyWs As Worksheet
    Dim logWs As Worksheet
    Dim targetMonth As String
    Dim dailyFolder As String
    Dim dailyFiles As Collection
    Dim rowMap As Object
    Dim lastMonthlyRow As Long
    Dim fileIndex As Long
    Dim processedFiles As Long
    Dim importedRows As Long
    Dim skippedRows As Long
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim applicationStateChanged As Boolean

    On Error GoTo FatalError

    Set monthlyBook = ThisWorkbook
    Set logWs = PrepareLogSheet(monthlyBook)
    WriteLog logWs, "INFO", "開始", "月次クレーム集計の取込を開始しました。"

    Set monthlyWs = GetWorksheetOrNothing(monthlyBook, MONTHLY_SHEET_NAME)
    If monthlyWs Is Nothing Then
        WriteLog logWs, "ERROR", MONTHLY_SHEET_NAME, "月次集計シートが見つからないため終了しました。"
        GoTo CleanExit
    End If

    targetMonth = ExtractTargetMonth(monthlyBook.Name)
    If Len(targetMonth) = 0 Then
        WriteLog logWs, "ERROR", monthlyBook.Name, "ブック名から対象年月YYMMを取得できないため終了しました。"
        GoTo CleanExit
    End If
    WriteLog logWs, "INFO", "対象月", targetMonth

    If Len(monthlyBook.Path) = 0 Then
        WriteLog logWs, "ERROR", monthlyBook.Name, "月次ブックが保存されていないため日次フォルダを特定できません。"
        GoTo CleanExit
    End If

    dailyFolder = CombinePath(CombinePath(monthlyBook.Path, "daily"), targetMonth)
    Set dailyFiles = CollectDailyFiles(dailyFolder, targetMonth, logWs)
    If dailyFiles.Count = 0 Then
        WriteLog logWs, "ERROR", dailyFolder, "対象の日次ファイルが見つからないため終了しました。"
        GoTo CleanExit
    End If

    lastMonthlyRow = LastUsedRow(monthlyWs, COL_BRANCH_CODE)
    If lastMonthlyRow < FIRST_DATA_ROW Then
        WriteLog logWs, "ERROR", MONTHLY_SHEET_NAME, "月次集計シートに転記先データ行がありません。"
        GoTo CleanExit
    End If

    Set rowMap = BuildMonthlyRowMap(monthlyWs, lastMonthlyRow, logWs)
    If rowMap.Count = 0 Then
        WriteLog logWs, "ERROR", MONTHLY_SHEET_NAME, "支店コード+業務コードの転記先を作成できませんでした。"
        GoTo CleanExit
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    applicationStateChanged = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ClearAndInitializeMonthlyValues monthlyWs, lastMonthlyRow
    WriteLog logWs, "INFO", MONTHLY_SHEET_NAME, "転記前に日別列E:AIと月合計AJをクリアし、0で初期化しました。"

    For fileIndex = 1 To dailyFiles.Count
        ProcessDailyWorkbook CStr(dailyFiles(fileIndex)), targetMonth, monthlyWs, rowMap, logWs, importedRows, skippedRows
        processedFiles = processedFiles + 1
    Next fileIndex

    CalculateMonthlyTotals monthlyWs, lastMonthlyRow
    WriteLog logWs, "INFO", MONTHLY_SHEET_NAME, "月合計AJを再計算しました。"
    WriteLog logWs, "INFO", "終了", "処理ファイル数=" & processedFiles & _
        "、転記行数=" & importedRows & "、スキップ行数=" & skippedRows

CleanExit:
    On Error Resume Next
    If applicationStateChanged Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.DisplayAlerts = oldDisplayAlerts
        Application.Calculation = oldCalculation
    End If
    If Not logWs Is Nothing Then
        logWs.Columns("A:D").AutoFit
    End If
    On Error GoTo 0
    Exit Sub

FatalError:
    On Error Resume Next
    If Not logWs Is Nothing Then
        WriteLog logWs, "ERROR", "致命的エラー", Err.Number & ": " & Err.Description
    End If
    Resume CleanExit
End Sub

Private Sub ProcessDailyWorkbook(ByVal dailyFilePath As String, ByVal targetMonth As String, _
                                 ByVal monthlyWs As Worksheet, ByVal rowMap As Object, _
                                 ByVal logWs As Worksheet, ByRef importedRows As Long, _
                                 ByRef skippedRows As Long)
    Dim dailyBook As Workbook
    Dim dailyWs As Worksheet
    Dim dayNumber As Long
    Dim dayCol As Long
    Dim lastDailyRow As Long
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim rowKey As String
    Dim rawCount As Variant
    Dim claimCount As Double
    Dim monthlyRow As Long
    Dim fileImportedRows As Long
    Dim fileSkippedRows As Long

    On Error GoTo FileError

    dayNumber = GetDayNumberFromDailyFile(dailyFilePath, targetMonth)
    If dayNumber < 1 Or dayNumber > 31 Then
        WriteLog logWs, "WARN", GetFileNameOnly(dailyFilePath), "ファイル名から有効な日付を取得できないためスキップしました。"
        skippedRows = skippedRows + 1
        Exit Sub
    End If
    dayCol = DAY_START_COL + dayNumber - 1

    Set dailyBook = Workbooks.Open(Filename:=dailyFilePath, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
    Set dailyWs = GetWorksheetOrNothing(dailyBook, DAILY_SHEET_NAME)
    If dailyWs Is Nothing Then
        WriteLog logWs, "WARN", GetFileNameOnly(dailyFilePath), "日次集計シートが見つからないためファイルをスキップしました。"
        skippedRows = skippedRows + 1
        GoTo CloseDailyBook
    End If

    lastDailyRow = LastUsedRow(dailyWs, COL_BRANCH_CODE)
    If lastDailyRow < FIRST_DATA_ROW Then
        WriteLog logWs, "WARN", GetFileNameOnly(dailyFilePath), "日次集計シートにデータ行がないためファイルをスキップしました。"
        skippedRows = skippedRows + 1
        GoTo CloseDailyBook
    End If

    For rowIndex = FIRST_DATA_ROW To lastDailyRow
        branchCode = NormalizeKeyPart(dailyWs.Cells(rowIndex, COL_BRANCH_CODE).Value)
        businessCode = NormalizeKeyPart(dailyWs.Cells(rowIndex, COL_BUSINESS_CODE).Value)

        If Len(branchCode) = 0 And Len(businessCode) = 0 Then
            GoTo ContinueRow
        End If

        If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
            fileSkippedRows = fileSkippedRows + 1
            WriteLog logWs, "WARN", GetFileNameOnly(dailyFilePath) & "!" & rowIndex, "支店コードまたは業務コードが空のためスキップしました。"
            GoTo ContinueRow
        End If

        rawCount = dailyWs.Cells(rowIndex, COL_CLAIM_COUNT).Value
        If Not HasNumericValue(rawCount) Then
            fileSkippedRows = fileSkippedRows + 1
            WriteLog logWs, "WARN", GetFileNameOnly(dailyFilePath) & "!" & rowIndex, "クレーム件数が数値ではないためスキップしました。"
            GoTo ContinueRow
        End If

        rowKey = MakeKey(branchCode, businessCode)
        If Not rowMap.Exists(rowKey) Then
            fileSkippedRows = fileSkippedRows + 1
            WriteLog logWs, "WARN", GetFileNameOnly(dailyFilePath) & "!" & rowIndex, _
                "月次集計に一致する支店コード+業務コードがないためスキップしました。キー=" & rowKey
            GoTo ContinueRow
        End If

        claimCount = CDbl(rawCount)
        monthlyRow = CLng(rowMap(rowKey))
        monthlyWs.Cells(monthlyRow, dayCol).Value = NzNumber(monthlyWs.Cells(monthlyRow, dayCol).Value) + claimCount
        fileImportedRows = fileImportedRows + 1

ContinueRow:
    Next rowIndex

CloseDailyBook:
    If Not dailyBook Is Nothing Then
        dailyBook.Close SaveChanges:=False
    End If
    importedRows = importedRows + fileImportedRows
    skippedRows = skippedRows + fileSkippedRows
    WriteLog logWs, "INFO", GetFileNameOnly(dailyFilePath), _
        "取込行数=" & fileImportedRows & "、スキップ行数=" & fileSkippedRows
    Exit Sub

FileError:
    On Error Resume Next
    If Not dailyBook Is Nothing Then
        dailyBook.Close SaveChanges:=False
    End If
    skippedRows = skippedRows + 1
    WriteLog logWs, "ERROR", GetFileNameOnly(dailyFilePath), Err.Number & ": " & Err.Description
End Sub

Private Function PrepareLogSheet(ByVal targetBook As Workbook) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheetOrNothing(targetBook, LOG_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
    End If

    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("日時", "レベル", "対象", "内容")
    ws.Rows(1).Font.Bold = True
    Set PrepareLogSheet = ws
End Function

Private Sub WriteLog(ByVal logWs As Worksheet, ByVal levelText As String, ByVal targetText As String, ByVal messageText As String)
    Dim nextRow As Long

    If logWs Is Nothing Then Exit Sub

    nextRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    logWs.Cells(nextRow, 1).Value = Now
    logWs.Cells(nextRow, 2).Value = levelText
    logWs.Cells(nextRow, 3).Value = targetText
    logWs.Cells(nextRow, 4).Value = messageText
End Sub

Private Function GetWorksheetOrNothing(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetOrNothing = targetBook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function CollectDailyFiles(ByVal dailyFolder As String, ByVal targetMonth As String, ByVal logWs As Worksheet) As Collection
    Dim result As Collection
    Dim fileName As String
    Dim searchPattern As String

    Set result = New Collection

    If Not FolderExists(dailyFolder) Then
        WriteLog logWs, "ERROR", dailyFolder, "日次フォルダが見つかりません。"
        Set CollectDailyFiles = result
        Exit Function
    End If

    searchPattern = CombinePath(dailyFolder, "クレーム集計" & targetMonth & "??.xlsx")
    fileName = Dir(searchPattern, vbNormal)
    Do While Len(fileName) > 0
        result.Add CombinePath(dailyFolder, fileName)
        fileName = Dir()
    Loop

    WriteLog logWs, "INFO", dailyFolder, "検出した日次ファイル数=" & result.Count
    Set CollectDailyFiles = result
End Function

Private Function BuildMonthlyRowMap(ByVal monthlyWs As Worksheet, ByVal lastMonthlyRow As Long, ByVal logWs As Worksheet) As Object
    Dim result As Object
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim rowKey As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For rowIndex = FIRST_DATA_ROW To lastMonthlyRow
        branchCode = NormalizeKeyPart(monthlyWs.Cells(rowIndex, COL_BRANCH_CODE).Value)
        businessCode = NormalizeKeyPart(monthlyWs.Cells(rowIndex, COL_BUSINESS_CODE).Value)

        If Len(branchCode) = 0 And Len(businessCode) = 0 Then
            GoTo ContinueRow
        End If

        If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
            WriteLog logWs, "WARN", MONTHLY_SHEET_NAME & "!" & rowIndex, "支店コードまたは業務コードが空の月次行です。"
            GoTo ContinueRow
        End If

        rowKey = MakeKey(branchCode, businessCode)
        If result.Exists(rowKey) Then
            WriteLog logWs, "WARN", MONTHLY_SHEET_NAME & "!" & rowIndex, _
                "支店コード+業務コードが重複しているため最初の行を使用します。キー=" & rowKey
        Else
            result.Add rowKey, rowIndex
        End If

ContinueRow:
    Next rowIndex

    Set BuildMonthlyRowMap = result
End Function

Private Sub ClearAndInitializeMonthlyValues(ByVal monthlyWs As Worksheet, ByVal lastMonthlyRow As Long)
    monthlyWs.Range(monthlyWs.Cells(FIRST_DATA_ROW, DAY_START_COL), monthlyWs.Cells(lastMonthlyRow, TOTAL_COL)).ClearContents
    monthlyWs.Range(monthlyWs.Cells(FIRST_DATA_ROW, DAY_START_COL), monthlyWs.Cells(lastMonthlyRow, DAY_END_COL)).Value = 0
    monthlyWs.Range(monthlyWs.Cells(FIRST_DATA_ROW, TOTAL_COL), monthlyWs.Cells(lastMonthlyRow, TOTAL_COL)).Value = 0
End Sub

Private Sub CalculateMonthlyTotals(ByVal monthlyWs As Worksheet, ByVal lastMonthlyRow As Long)
    Dim rowIndex As Long

    For rowIndex = FIRST_DATA_ROW To lastMonthlyRow
        monthlyWs.Cells(rowIndex, TOTAL_COL).Value = Application.WorksheetFunction.Sum( _
            monthlyWs.Range(monthlyWs.Cells(rowIndex, DAY_START_COL), monthlyWs.Cells(rowIndex, DAY_END_COL)))
    Next rowIndex
End Sub

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    Dim candidateRow As Long

    candidateRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
    If candidateRow = 1 And Len(Trim$(CStr(ws.Cells(1, columnIndex).Value))) = 0 Then
        LastUsedRow = 0
    Else
        LastUsedRow = candidateRow
    End If
End Function

Private Function ExtractTargetMonth(ByVal workbookName As String) As String
    Dim baseName As String
    Dim dotPosition As Long
    Dim index As Long
    Dim token As String

    dotPosition = InStrRev(workbookName, ".")
    If dotPosition > 1 Then
        baseName = Left$(workbookName, dotPosition - 1)
    Else
        baseName = workbookName
    End If

    For index = Len(baseName) - 3 To 1 Step -1
        token = Mid$(baseName, index, 4)
        If IsAllDigits(token) Then
            ExtractTargetMonth = token
            Exit Function
        End If
    Next index
End Function

Private Function GetDayNumberFromDailyFile(ByVal dailyFilePath As String, ByVal targetMonth As String) As Long
    Dim fileName As String
    Dim dateToken As String

    fileName = GetFileNameOnly(dailyFilePath)
    dateToken = ExtractDailyDateToken(fileName, targetMonth)
    If Len(dateToken) = 6 Then
        GetDayNumberFromDailyFile = CLng(Right$(dateToken, 2))
    Else
        GetDayNumberFromDailyFile = 0
    End If
End Function

Private Function ExtractDailyDateToken(ByVal fileName As String, ByVal targetMonth As String) As String
    Dim index As Long
    Dim token As String

    For index = 1 To Len(fileName) - 5
        token = Mid$(fileName, index, 6)
        If IsAllDigits(token) And Left$(token, 4) = targetMonth Then
            ExtractDailyDateToken = token
            Exit Function
        End If
    Next index
End Function

Private Function IsAllDigits(ByVal textValue As String) As Boolean
    Dim index As Long
    Dim oneChar As String

    If Len(textValue) = 0 Then Exit Function

    For index = 1 To Len(textValue)
        oneChar = Mid$(textValue, index, 1)
        If oneChar < "0" Or oneChar > "9" Then
            Exit Function
        End If
    Next index

    IsAllDigits = True
End Function

Private Function MakeKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeKey = NormalizeKeyPart(branchCode) & KEY_DELIMITER & NormalizeKeyPart(businessCode)
End Function

Private Function NormalizeKeyPart(ByVal value As Variant) As String
    On Error GoTo InvalidValue

    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function

    NormalizeKeyPart = Trim$(CStr(value))
    Exit Function

InvalidValue:
    NormalizeKeyPart = vbNullString
End Function

Private Function HasNumericValue(ByVal value As Variant) As Boolean
    On Error GoTo InvalidValue

    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function

    HasNumericValue = IsNumeric(value)
    Exit Function

InvalidValue:
    HasNumericValue = False
End Function

Private Function NzNumber(ByVal value As Variant) As Double
    On Error GoTo InvalidValue

    If HasNumericValue(value) Then
        NzNumber = CDbl(value)
    End If
    Exit Function

InvalidValue:
    NzNumber = 0
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo NotFound
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    Exit Function

NotFound:
    FolderExists = False
End Function

Private Function CombinePath(ByVal folderPath As String, ByVal childName As String) As String
    If Len(folderPath) = 0 Then
        CombinePath = childName
    ElseIf Right$(folderPath, 1) = Application.PathSeparator Then
        CombinePath = folderPath & childName
    Else
        CombinePath = folderPath & Application.PathSeparator & childName
    End If
End Function

Private Function GetFileNameOnly(ByVal filePath As String) As String
    Dim separatorPosition As Long

    separatorPosition = InStrRev(filePath, Application.PathSeparator)
    If separatorPosition = 0 Then separatorPosition = InStrRev(filePath, "\")
    If separatorPosition = 0 Then separatorPosition = InStrRev(filePath, "/")

    If separatorPosition > 0 Then
        GetFileNameOnly = Mid$(filePath, separatorPosition + 1)
    Else
        GetFileNameOnly = filePath
    End If
End Function
```
