# codex-multi-agent-review 修正前検証

## 前提

- 対象: `outputs/codex-multi-agent-review/first-output.md`
- 検証方法: `docs/task-spec.md` と `prompts/common-task-prompt.md` に対する静的確認
- Excel実行: Issue #5では未実施

## 見つかった主な欠陥

1. `BuildMasterSets` / `FindHeaderColumn` がマスタ見出しを1行目だけから探し、見つからない場合は致命エラーで停止する。Issue #3のサンプルマスタは支店マスタ2行目、業務マスタ12行目、対応表24行目の区画構造であり、仕様では見出し欠落時も月次 `DETAIL` 行からフォールバックすべき。
2. `ResetMonthlyArea` が `DETAIL` と `SUBTOTAL` の `H:AM` を実行開始時にクリアしている。仕様上の事前クリア範囲は `DETAIL` 行の `H:AL` 日別転記セルだけで、`AM` 月合計や `SUBTOTAL` 行は最後の再計算で更新するほうが安全。
3. `ProcessDailyRows` は `処理区分 <> "ok"` の行をログなしでスキップする。仕様では `ok` 以外の行も異常系サンプルとして転記せずログ対象にする。
4. `RecalculateMonthlyTotals` が0件の `DETAIL` 月合計、`SUBTOTAL` 日別、`SUBTOTAL` 月合計に `0` を書き込む。期待結果が空欄を前提にするセルと一致しない可能性がある。
5. ログに `ThisWorkbook.FullName`、`dailyRoot`、`filePath` をフルパスのまま記録している。公開実験用ログとしては、`ThisWorkbook.Path` からの相対パス化が必要。
6. 月次ブック未保存時の明示チェックがない。`ThisWorkbook.Path` が空の場合、探索起点やログが分かりにくい状態になる。
7. ログ列にクレーム件数の独立列がなく、非数値や空欄などの元値がメッセージに埋もれる。Issue #5の比較ログとしては、問題行の件数値を構造化して残せるほうが望ましい。

## 未実行項目

- VBAコンパイル確認、Excelでの実行、期待結果CSVとの突合はIssue #6以降で実施する。
