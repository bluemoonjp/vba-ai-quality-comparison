# サンプル

初回検証用の合成サンプルです。支店、業務、件数はすべて架空です。

## 内容

- `source/branches.csv`: 支店マスタ。
- `source/business-lines.csv`: 業務マスタ。
- `source/daily-records.csv`: 日次データのCSV正本。
- `expected/monthly-expected.csv`: 期待結果。
- `workbooks/`: 公開チェック済みサンプル `.xlsx`。

## 公開チェック

サンプルworkbookは合成CSVから生成し、実データ、実パス、secretを含めません。モデル成果物の `.xlsm` は保存しません。
