# first output: codex-chatgpt-medium

- 条件ID: `codex-chatgpt-medium`
- モデル/推論設定として記録する条件: reasoning medium
- 実行形態: ChatGPT代替の prompt応答のみ

## 回答

以下のコードを、月次ファイル `月次クレーム集計YYMM.xlsm` の標準モジュールに貼り付けてください。

- 入口プロシージャ名: `ImportDailyClaimsToMonthly`
- コードの配置場所: 月次ファイル側の標準モジュール。例: `modClaimsMonthlyImport`
- 実行方法: 月次ファイルを開き、VBEで標準モジュールへ貼り付けたうえで、Excelのマクロ一覧またはVBEから `ImportDailyClaimsToMonthly` を実行します。
- 日次ファイルの場所: 月次ファイルと同じフォルダ配下の `daily\YYMM\` を見に行きます。`YYMM` は月次ファイル名内の4桁数字から取得します。
- 参照設定: 追加参照は不要です。`Scripting.Dictionary` は late binding で使います。
- ログ出力方法: `処理ログ` シートを作成または初期化し、処理結果・警告・エラーを一覧で残します。MsgBoxでは停止しません。
- エラー処理: 想定外エラーも `処理ログ` に記録し、Excelの表示設定を復元して終了します。日次ファイル単位のエラーはログに残し、可能な範囲で処理を継続します。
- 再実行安全性: 転記前に `月次集計` シートの E:AI の日別列と AJ の月合計をクリアし、日別列を0で初期化してから日次ファイルを読み直します。そのため、同じブックで2回実行しても前回分に二重加算されません。

```vb
Option Explicit

Private Const SHEET_MONTHLY As String = "月次集計"
Private Const SHEET_DAILY As String = "日次集計"
Private Const SHEET_LOG As String = "処理ログ"

Private Const COL_BRANCH_CODE As Long = 1
Private Const COL_BUSINESS_CODE As Long = 3
Private Const COL_DAY_FIRST As Long = 5   ' E列: 1日
Private Const COL_DAY_LAST As Long = 35   ' AI列: 31日
Private Const COL_MONTH_TOTAL As Long = 36 ' AJ列: 月合計

Public Sub ImportDailyClaimsToMonthly()
    Dim monthlyWb As Workbook
    Dim monthlyWs As Worksheet
    Dim logWs As Worksheet
    Dim monthlyMap As Object
    Dim targetMonth As String
    Dim dailyFolder As String
    Dim fileName As String
    Dim lastMonthlyRow As Long
    Dim processedFiles As Long
    Dim transferredRows As Long
    Dim skippedRows As Long
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean

    On Error GoTo FatalError

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.StatusBar = "クレーム日次集計を転記しています..."

    Set monthlyWb = ThisWorkbook
    Set logWs = PrepareLogSheet(monthlyWb)
    WriteLog logWs, "INFO", "処理開始", monthlyWb.Name

    Set monthlyWs = GetWorksheet(monthlyWb, SHEET_MONTHLY)
    If monthlyWs Is Nothing Then
        WriteLog logWs, "ERROR", "月次集計シートが見つかりません。", SHEET_MONTHLY
        GoTo CleanExit
    End If

    targetMonth = ExtractTargetMonth(monthlyWb.Name)
    If Len(targetMonth) <> 4 Then
        WriteLog logWs, "ERROR", "月次ファイル名から対象月YYMMを取得できません。", monthlyWb.Name
        GoTo CleanExit
    End If

    If Len(monthlyWb.Path) = 0 Then
        WriteLog logWs, "ERROR", "月次ファイルが保存されていないため、日次フォルダを判断できません。", monthlyWb.Name
        GoTo CleanExit
    End If

    dailyFolder = monthlyWb.Path & Application.PathSeparator & "daily" & Application.PathSeparator & targetMonth
    If Not FolderExists(dailyFolder) Then
        WriteLog logWs, "ERROR", "日次フォルダが見つかりません。", dailyFolder
        GoTo CleanExit
    End If

    lastMonthlyRow = monthlyWs.Cells(monthlyWs.Rows.Count, COL_BRANCH_CODE).End(xlUp).Row
    If lastMonthlyRow < 2 Then
        WriteLog logWs, "ERROR", "月次集計シートに転記対象行がありません。", SHEET_MONTHLY
        GoTo CleanExit
    End If

    ClearMonthlyValues monthlyWs, lastMonthlyRow
    Set monthlyMap = BuildMonthlyRowMap(monthlyWs, lastMonthlyRow, logWs)

    fileName = Dir(dailyFolder & Application.PathSeparator & "クレーム集計" & targetMonth & "*.xls*")
    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            ProcessDailyWorkbook dailyFolder & Application.PathSeparator & fileName, _
                                 fileName, _
                                 targetMonth, _
                                 monthlyWs, _
                                 monthlyMap, _
                                 logWs, _
                                 processedFiles, _
                                 transferredRows, _
                                 skippedRows
        End If
        fileName = Dir()
    Loop

    CalculateMonthlyTotals monthlyWs, lastMonthlyRow

    If processedFiles = 0 Then
        WriteLog logWs, "WARN", "処理対象の日次ファイルがありませんでした。", dailyFolder
    End If

    WriteLog logWs, "INFO", "処理完了", _
             "対象月=" & targetMonth & _
             ", 処理ファイル数=" & processedFiles & _
             ", 転記行数=" & transferredRows & _
             ", スキップ行数=" & skippedRows

CleanExit:
    Application.StatusBar = False
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

FatalError:
    If Not logWs Is Nothing Then
        WriteLog logWs, "ERROR", "予期しないエラー: " & Err.Number & " " & Err.Description, "ImportDailyClaimsToMonthly"
    End If
    Resume CleanExit
End Sub

Private Function PrepareLogSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheet(wb, SHEET_LOG)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_LOG
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1:D1").Value = Array("日時", "レベル", "内容", "詳細")
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:D").ColumnWidth = 24

    Set PrepareLogSheet = ws
End Function

Private Sub WriteLog(ByVal logWs As Worksheet, ByVal levelText As String, ByVal messageText As String, ByVal detailText As String)
    Dim nextRow As Long

    If logWs Is Nothing Then Exit Sub

    nextRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    logWs.Cells(nextRow, 1).Value = Now
    logWs.Cells(nextRow, 2).Value = levelText
    logWs.Cells(nextRow, 3).Value = messageText
    logWs.Cells(nextRow, 4).Value = detailText
End Sub

Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function ExtractTargetMonth(ByVal workbookName As String) As String
    Dim i As Long
    Dim ch As String
    Dim digits As String

    For i = 1 To Len(workbookName)
        ch = Mid$(workbookName, i, 1)
        If ch Like "#" Then
            digits = digits & ch
            If Len(digits) >= 4 Then
                ExtractTargetMonth = Left$(digits, 4)
                Exit Function
            End If
        Else
            digits = vbNullString
        End If
    Next i
End Function

Private Function ExtractDailyDate(ByVal fileName As String) As String
    Dim i As Long
    Dim ch As String
    Dim digits As String

    For i = 1 To Len(fileName)
        ch = Mid$(fileName, i, 1)
        If ch Like "#" Then
            digits = digits & ch
            If Len(digits) >= 6 Then
                ExtractDailyDate = Right$(digits, 6)
            End If
        Else
            digits = vbNullString
        End If
    Next i
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = (Len(Dir(folderPath, vbDirectory)) > 0)
End Function

Private Sub ClearMonthlyValues(ByVal monthlyWs As Worksheet, ByVal lastRow As Long)
    With monthlyWs
        .Range(.Cells(2, COL_DAY_FIRST), .Cells(lastRow, COL_MONTH_TOTAL)).ClearContents
        .Range(.Cells(2, COL_DAY_FIRST), .Cells(lastRow, COL_DAY_LAST)).Value = 0
        .Range(.Cells(2, COL_MONTH_TOTAL), .Cells(lastRow, COL_MONTH_TOTAL)).Value = 0
    End With
End Sub

Private Function BuildMonthlyRowMap(ByVal monthlyWs As Worksheet, ByVal lastRow As Long, ByVal logWs As Worksheet) As Object
    Dim rowMap As Object
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim key As String

    Set rowMap = CreateObject("Scripting.Dictionary")

    For rowIndex = 2 To lastRow
        branchCode = Trim$(CStr(monthlyWs.Cells(rowIndex, COL_BRANCH_CODE).Value))
        businessCode = Trim$(CStr(monthlyWs.Cells(rowIndex, COL_BUSINESS_CODE).Value))

        If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
            WriteLog logWs, "WARN", "月次集計のキー項目が空のため行を無視しました。", "行=" & rowIndex
        Else
            key = MakeKey(branchCode, businessCode)
            If rowMap.Exists(key) Then
                WriteLog logWs, "WARN", "月次集計に同じ支店コード+業務コードの行があります。先の行へ転記します。", _
                         "キー=" & key & ", 後続行=" & rowIndex
            Else
                rowMap.Add key, rowIndex
            End If
        End If
    Next rowIndex

    Set BuildMonthlyRowMap = rowMap
End Function

Private Sub ProcessDailyWorkbook(ByVal filePath As String, _
                                 ByVal fileName As String, _
                                 ByVal targetMonth As String, _
                                 ByVal monthlyWs As Worksheet, _
                                 ByVal monthlyMap As Object, _
                                 ByVal logWs As Worksheet, _
                                 ByRef processedFiles As Long, _
                                 ByRef transferredRows As Long, _
                                 ByRef skippedRows As Long)
    Dim dailyWb As Workbook
    Dim dailyWs As Worksheet
    Dim dailyDate As String
    Dim dayNumber As Long
    Dim dayColumn As Long
    Dim lastDailyRow As Long
    Dim rowIndex As Long
    Dim branchCode As String
    Dim businessCode As String
    Dim claimValue As Variant
    Dim claimCount As Double
    Dim key As String
    Dim monthlyRow As Long

    On Error GoTo DailyError

    dailyDate = ExtractDailyDate(fileName)
    If Len(dailyDate) <> 6 Then
        WriteLog logWs, "WARN", "日次ファイル名からYYMMDDを取得できないためスキップしました。", fileName
        Exit Sub
    End If

    If Left$(dailyDate, 4) <> targetMonth Then
        WriteLog logWs, "WARN", "対象月と異なる日次ファイルをスキップしました。", fileName
        Exit Sub
    End If

    dayNumber = CLng(Right$(dailyDate, 2))
    If dayNumber < 1 Or dayNumber > 31 Then
        WriteLog logWs, "WARN", "日付の日が1から31の範囲外のためスキップしました。", fileName
        Exit Sub
    End If
    dayColumn = COL_DAY_FIRST + dayNumber - 1

    Set dailyWb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set dailyWs = GetWorksheet(dailyWb, SHEET_DAILY)
    If dailyWs Is Nothing Then
        WriteLog logWs, "WARN", "日次集計シートが見つからないためファイルをスキップしました。", fileName
        GoTo DailyCleanExit
    End If

    lastDailyRow = dailyWs.Cells(dailyWs.Rows.Count, 1).End(xlUp).Row
    If lastDailyRow < 2 Then
        WriteLog logWs, "WARN", "日次集計シートにデータ行がありません。", fileName
        processedFiles = processedFiles + 1
        GoTo DailyCleanExit
    End If

    For rowIndex = 2 To lastDailyRow
        branchCode = Trim$(CStr(dailyWs.Cells(rowIndex, 1).Value))
        businessCode = Trim$(CStr(dailyWs.Cells(rowIndex, 3).Value))
        claimValue = dailyWs.Cells(rowIndex, 5).Value

        If Len(branchCode) = 0 Or Len(businessCode) = 0 Then
            skippedRows = skippedRows + 1
            WriteLog logWs, "WARN", "日次行のキー項目が空のためスキップしました。", fileName & " 行=" & rowIndex
        ElseIf Not IsNumeric(claimValue) Then
            skippedRows = skippedRows + 1
            WriteLog logWs, "WARN", "クレーム件数が数値ではないためスキップしました。", fileName & " 行=" & rowIndex
        Else
            key = MakeKey(branchCode, businessCode)
            If monthlyMap.Exists(key) Then
                claimCount = CDbl(claimValue)
                monthlyRow = CLng(monthlyMap(key))
                monthlyWs.Cells(monthlyRow, dayColumn).Value = NzNumber(monthlyWs.Cells(monthlyRow, dayColumn).Value) + claimCount
                transferredRows = transferredRows + 1
            Else
                skippedRows = skippedRows + 1
                WriteLog logWs, "WARN", "月次集計に該当する支店コード+業務コードがないためスキップしました。", _
                         fileName & " 行=" & rowIndex & " キー=" & key
            End If
        End If
    Next rowIndex

    processedFiles = processedFiles + 1
    WriteLog logWs, "INFO", "日次ファイルを処理しました。", fileName

DailyCleanExit:
    On Error Resume Next
    If Not dailyWb Is Nothing Then dailyWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub

DailyError:
    WriteLog logWs, "ERROR", "日次ファイル処理中のエラー: " & Err.Number & " " & Err.Description, fileName
    Resume DailyCleanExit
End Sub

Private Function MakeKey(ByVal branchCode As String, ByVal businessCode As String) As String
    MakeKey = branchCode & "|" & businessCode
End Function

Private Function NzNumber(ByVal valueToCheck As Variant) As Double
    If IsError(valueToCheck) Or IsEmpty(valueToCheck) Or Len(Trim$(CStr(valueToCheck))) = 0 Then
        NzNumber = 0
    ElseIf IsNumeric(valueToCheck) Then
        NzNumber = CDbl(valueToCheck)
    Else
        NzNumber = 0
    End If
End Function

Private Sub CalculateMonthlyTotals(ByVal monthlyWs As Worksheet, ByVal lastRow As Long)
    Dim rowIndex As Long

    For rowIndex = 2 To lastRow
        monthlyWs.Cells(rowIndex, COL_MONTH_TOTAL).Value = _
            Application.WorksheetFunction.Sum(monthlyWs.Range(monthlyWs.Cells(rowIndex, COL_DAY_FIRST), monthlyWs.Cells(rowIndex, COL_DAY_LAST)))
    Next rowIndex
End Sub
```

## 補足

- `.bas` として保存する場合は、上記コードブロック内の `Option Explicit` から最後の `End Sub` までを標準モジュール用テキストとして保存できます。
- `処理ログ` シートは実行時に作成または初期化します。ログ保存のためのシート追加であり、日次ファイルや月次集計表の列構造変更を前提にしていません。
- 日次ファイルは読み取り専用で開き、保存せずに閉じます。
- 月次表に存在しない支店コード+業務コードは追加せず、ログへ記録してスキップします。
