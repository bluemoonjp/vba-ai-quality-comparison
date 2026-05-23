# 日本語化Audit Issue #8

## 目的

Issue #8では、人が読むドキュメント、prompt、template、CSV表示値、Excel workbook表示を日本語へ揃えました。

安定IDは比較・再生成・評価のためASCIIのまま維持しています。

## 日本語化した範囲

- `README.md`, `AGENTS.md`
- `docs/*.md`
- `prompts/*.md`
- `outputs/*.md`
- `tests/*.md`
- `samples/*.md`
- `src/README.md`
- `samples/source/*.csv` の表示値
- `samples/expected/*.csv` の表示値
- `samples/checked/` 配下の `.xlsx` 表示値
- `tools/build-samples.mjs` が生成する見出し、ラベル、備考

## 意図的にASCIIのまま残したもの

- `LICENSE` のMIT原文
- GitHub URL
- `condition id`
- `branch_code`, `business_code`, `status`
- `DETAIL`, `SUBTOTAL`
- CSVヘッダー
- ファイル名パターン
- VBA評価用の機械的識別子

## 生成結果

`tools/build-samples.mjs --force` で次を再生成しました。

- 日次workbook: 11件
- 月次workbook: 3件
- 期待結果CSV: 2件

`.xlsm` は引き続き作成していません。VBA生成・埋め込み後のIssueで扱います。

## 検証メモ

- `git diff --check` でMarkdown/CSV/JS差分の空白エラーがないことを確認しました。
- 主要な旧英語表示値の `rg` scanで該当なしを確認しました。
- 月次3ファイルの `月次集計`, `期待結果`, `マスタ`, `異常系` 主要rangeをinspectし、日本語表示を確認しました。
- 月次3ファイルと代表日次ファイルでformula/error scanを行い、`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A` がないことを確認しました。
- Officeバイナリは `samples/checked/` 配下の `.xlsx` 14件だけです。
- renderは月次3ファイルの全主要sheetで処理されましたが、artifact-tool側が最後に非0終了を返したため、inspect結果を主要証跡とします。

## 後続Issueへの引き渡し

Issue #5以降では、日本語化済みの `docs/task-spec.md` と `prompts/common-task-prompt.md` を正本として使います。
