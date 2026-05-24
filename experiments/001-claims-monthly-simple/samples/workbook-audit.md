# workbook公開チェック

対象:

- `samples/workbooks/月次クレーム集計2605.xlsx`
- `samples/workbooks/daily/2605/クレーム集計260501.xlsx`
- `samples/workbooks/daily/2605/クレーム集計260502.xlsx`
- `samples/workbooks/daily/2605/クレーム集計260515.xlsx`

## 生成元

`samples/source/*.csv` から `samples/source/build-sample-workbooks.mjs` で生成しました。データはすべて架空です。

## 確認結果

- artifact-tool inspectで月次 `月次集計!A1:F5` と日次 `日次集計!A1:E4` の主要セルを確認済み。
- `rg -a` でローカルユーザー名、クラウド同期パス、Windowsユーザーパス、secret、tokenを示す文字列がサンプルworkbook内に出ないことを確認済み。
- zip内 `docProps/*` は確認対象ファイルでは見つかりませんでした。
- renderは月次 `月次集計`, `マスタ`, `期待結果` の各シートでpreview byteを取得できましたが、Nodeプロセス終了コードは1でした。内容inspectは成功しているため、視覚確認は次回のExcel検証Issueで再確認します。

## 保存判断

公開チェック済みサンプル `.xlsx` として `samples/workbooks/` に保存します。モデル成果物としての `.xlsm` は保存しません。

## Issue #12 検収

2026-05-24 に、初回検証サンプルとして次を再確認しました。

- CSV正本は `samples/source/branches.csv`, `samples/source/business-lines.csv`, `samples/source/daily-records.csv` の3件です。
- 期待結果は `samples/expected/monthly-expected.csv` です。
- 公開チェック済みworkbookは、月次1件と日次3件の `.xlsx` です。
- `daily-records.csv` を支店コード + 業務コード + 日付で再集計し、`monthly-expected.csv` の `day_01`, `day_02`, `day_15`, `month_total` と一致することを確認しました。
- artifact-tool inspectで、月次workbookに `月次集計`, `マスタ`, `期待結果` があり、日次workbook 3件に `日次集計` があることを確認しました。
- 月次 `月次集計` は支店3件 x 業務3件の9行と、1日から31日、月合計の列を持ちます。
- tracked Office binaryは `samples/workbooks/` 配下の `.xlsx` 4件だけです。`.xlsm`, `.xls`, `.xlsb` はtrackedされていません。
- `rg -a` によるsecret/local path scanでは、実パス、ユーザー名、token、private source materialは検出されず、公開境界に関する文言だけがヒットしました。

月合計は期待結果CSVで検証できます。再実行安全性は、課題仕様で転記前クリアを必須にし、検証計画で二重加算防止を確認項目にしているため、このサンプルで評価できます。
