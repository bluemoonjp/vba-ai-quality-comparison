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
