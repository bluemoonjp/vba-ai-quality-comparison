# Issue #3 サンプル監査

監査日: 2026-05-23

## 対象範囲

Issue #3 では、合成サンプル仕様、CSV正本、期待結果、チェック済み `.xlsx` workbookを作成しました。`.xlsm`、VBAモジュール、AI実行用prompt、AI生成物は作成していません。

## 固定した判断

- 対象月:
  - `2026-02`: 28日の月
  - `2026-04`: 30日の月
  - `2026-05`: 31日の月
- 正本:
  - `samples/source/*.csv`
  - `samples/expected/*.csv`
  - `docs/sample-data-spec.md`
- チェック済み `.xlsx` は `samples/checked/` 配下の再生成可能な成果物。
- `.xlsm` はVBA実装フェーズまで延期。

## 生成成果物

- 日次workbook: `samples/checked/daily/` 配下に11件。
- 月次workbook: `samples/checked/monthly/` 配下に3件。
- 期待値CSV:
  - `samples/expected/monthly-expected.csv`
  - `samples/expected/anomaly-expected.csv`

## 検証サマリ

- `tools/build-samples.mjs` でサンプルworkbookと期待値CSVを生成した。
- 生成後、月次workbookの主要な `期待結果` rangeをinspectした。
- formula/error scanで `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A` は検出されなかった。
- このローカル環境ではrender commandが表示出力後に非ゼロ終了を返す挙動があったため、安定したcloseout証跡はrange inspectとformula/error scanを正とした。
- Officeバイナリは `samples/checked/` 配下に限定されていた。
- secret/path scanのヒットは公開安全ルール文だけだった。

## 公開安全メモ

- 支店、地域、業務、フォルダ名はすべて架空。
- workbookは `.xlsx` であり、`.xlsm` ではない。
- commitしたworkbookはレイアウトと期待結果のサンプルであり、実運用データではない。

## 引き渡し

- Issue #4 は `docs/sample-data-spec.md`, `samples/source/*.csv`, `samples/expected/*.csv` を使ってAI向け課題仕様を作成する。
- Issue #5 では、AI出力収集中にこれらのサンプルを変更しない。
- Issue #6 では、チェック済み `.xlsx` と期待値CSVを検証に使う。
