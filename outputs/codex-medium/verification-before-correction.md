# codex-medium 修正前検証

## 前提

- 対象: `outputs/codex-medium/first-output.md`
- 検証方法: `docs/task-spec.md` と `prompts/common-task-prompt.md` に対する静的確認
- Excel実行: Issue #5では未実施

## 見つかった主な欠陥

1. 重複日付の日次ファイルがある場合、`validFilesByDate.Remove CStr(dateKey)` により該当日付を全件スキップしている。仕様では、同一 `YYMMDD` が複数ある場合は1件だけ処理し、重複ファイルをログに記録する。
2. `ClearMonthlyValues` が `SUBTOTAL` 行と `AM` 列も事前クリアし、さらに月外日列を `wsMonthly.Columns(c).ClearContents` で列全体クリアしている。ヘッダーや関係ない行まで壊す可能性があり、「DETAIL 行の H:AL の日別転記セルだけをクリア」「行ラベル、マスタ、備考、無関係シートは壊さない」に反する。
3. `AddLog` はログ見出しに「件数」列を用意しているが、引数にも代入処理にも件数値がなく、空欄/非数値/未知キーなどの問題行で実際の件数値を記録できない。
4. `LoadMasterCodes` / `FindHeaderColumn` はマスタ見出しを1行目だけから探す。Issue #3のサンプルマスタは区画見出しが2行目、12行目、24行目にあるため、支店/業務マスタを正しく読めない可能性が高い。
5. 実行時ログに `dailyRoot` や `filePath` のフルパスをそのまま書く設計になっている。公開実験用のサンプルでは、実パスがログに残らないよう相対パス化するほうが安全。
6. `ThisWorkbook.Path` が空、つまり月次ブック未保存の場合の明示チェックがない。探索起点が不明瞭なまま処理される可能性がある。

## 未実行項目

- VBAコンパイル確認、Excelでの実行、期待結果CSVとの突合はIssue #6以降で実施する。
