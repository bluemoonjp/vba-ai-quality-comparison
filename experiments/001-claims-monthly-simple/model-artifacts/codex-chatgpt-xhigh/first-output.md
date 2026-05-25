# first output: codex-chatgpt-xhigh

- 条件ID: `codex-chatgpt-xhigh`
- モデル/推論設定として記録する条件: reasoning xhigh
- 実行形態: ChatGPT代替の prompt応答のみ
- 確認質問: なし。標準回答集の「月次ファイルと同じフォルダ配下に `daily/2605/` がある」を採用。

## 概要

月次ブックの `月次集計` シートにある支店コード + 業務コードの行へ、同じフォルダ配下の `daily\<YYMM>\` にある日次ブックを読み込み、ファイル名の `YYMMDD` から該当日列へクレーム件数を転記します。

入口プロシージャ名は `ImportDailyClaimsToMonthly` です。

## コードの配置場所

月次ブック `月次クレーム集計YYMM.xlsm` の標準モジュールに配置してください。`.bas` として保存・インポートする場合は、下記コードをそのまま `modClaimsMonthlyImport.bas` として扱えます。VBEへ直接貼り付ける環境で `Attribute VB_Name` 行が受け付けられない場合は、先頭の `Attribute VB_Name = "modClaimsMonthlyImport"` の1行だけ除外してください。

## 実行方法

1. 月次ブックを `.xlsm` として開きます。
2. 月次ブックと同じフォルダに `daily\<YYMM>\` フォルダを置き、その中に `クレーム集計YYMMDD.xlsx` など、ファイル名に6桁の `YYMMDD` を含む日次ファイルを置きます。
3. VBAのマクロ一覧から `ImportDailyClaimsToMonthly` を実行します。
4. 処理結果は `処理ログ` シートで確認します。MsgBoxでは停止しません。

## 参照設定の要否

追加参照設定は不要です。辞書は `CreateObject("Scripting.Dictionary")` による late binding を使うため、`Microsoft Scripting Runtime` の参照追加は不要です。外部通信も行いません。

## ログ出力方法

`処理ログ` シートを作成または再利用し、実行開始時に内容をクリアしてから、日時、レベル、ファイル、行、メッセージを一覧で出力します。日次ファイル単位の処理完了、対象外キー、数値不正、ファイル処理エラー、全体サマリーを記録します。

## エラー処理

月次シートがない、対象月をファイル名から判断できない、日次フォルダがないなどの全体エラーは `処理ログ` に記録して終了します。個別の日次ファイルで開けない、`日次集計` シートがないなどのエラーはログに残し、可能な範囲で次のファイル処理へ進みます。いずれの場合も MsgBox は使いません。

## 再実行しても二重加算しない理由

処理開始後、月次集計表のデータ行について `E:AJ`、つまり1日から31日までの日別列と月合計列を `ClearContents` でクリアし、日別列を0で初期化してから日次ファイルを読み込みます。前回実行時の転記値を残したまま加算しないため、同じファイルで2回実行しても二重加算になりません。

## VBAコード

```vb
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
```

## 固定パス・外部通信・構造変更について

実ローカルパス、ユーザー名、OneDriveパス、社内パスは埋め込んでいません。日次フォルダは実行中の月次ブック `ThisWorkbook.Path` から相対的に判断します。外部通信は行いません。月次シートと日次シートの指定済み構造を前提に読み書きしますが、サンプルworkbook自体の構造変更を前提にした処理は含めていません。ログ出力用に `処理ログ` シートは作成または再利用します。
