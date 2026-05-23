# codex-xhigh 修正前検証

## 前提

- 対象: `outputs/codex-xhigh/first-output.md`
- 検証方法: `docs/task-spec.md` と `prompts/common-task-prompt.md` に対する静的確認
- Excel実行: Issue #5では未実施

## 見つかった主な欠陥

1. 同一日付の日次ファイルが複数ある場合、`paths.Count > 1` でその日付の全ファイルをスキップしている。仕様では、同一 `YYMMDD` が複数ある場合は1件だけ処理し、重複ファイルをログに記録する。
2. `ClearTransferArea` が事前処理で `DETAIL` 行の月合計と `SUBTOTAL` 行の日別・月合計もクリアしている。仕様上の事前クリア範囲は `DETAIL` 行の `H:AL` 日別転記セルだけであり、月合計と小計行は最後の再計算で更新するほうが安全。
3. `RecalculateMonthlyTotals` が0件の `DETAIL` 月合計、`SUBTOTAL` 日別、`SUBTOTAL` 月合計に `0` を書き込む。期待結果が空欄を前提にするセルと一致しない可能性がある。
4. ログへ `dailyRoot` や `filePath` をフルパスのまま記録している。公開実験用ログとしては、`ThisWorkbook.Path` からの相対パス化が必要。
5. `LoadMasterCodeSets` が `UsedRange` 上の全 `支店コード` / `業務コード` 見出しから下方向に読むため、支店マスタ、業務マスタ、支店・業務対応表の区画を意図せず混ぜる可能性がある。Issue #3のサンプルマスタは2行目、12行目、24行目の区画前提なので、その構造に合わせた読み取りが望ましい。

## 未実行項目

- VBAコンパイル確認、Excelでの実行、期待結果CSVとの突合はIssue #6以降で実施する。
