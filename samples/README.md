# サンプル

このディレクトリには、公開可能な合成Excelサンプルと期待結果を置きます。

サンプル一式には、次の内容を含めます。

- `クレーム集計YYMMDD.xlsx` の日次ブックレイアウト
- `月次クレーム集計YYMM.xlsm` を想定した月次ブックレイアウト
- 支店マスタと業務マスタ
- 大支店の単独小計と、小支店の地域別小計
- 日次転記、月合計、小計の期待結果
- 検証用の異常系ケース

ここへcommitするworkbookは、repository READMEと `docs/public-boundary.md` の公開前チェックを通過している必要があります。

## 正本CSV

- `source/branches.csv`: 架空の支店マスタと小計グループ。
- `source/business-lines.csv`: 架空の業務マスタ。
- `source/daily-records.csv`: 日次workbookサンプルを生成するための元データ。

AIへ渡す課題仕様の正本は `../docs/task-spec.md` です。全AI条件へ渡す共通promptは `../prompts/common-task-prompt.md` です。

## 期待結果

- `expected/monthly-expected.csv`: 有効な日次行から生成される月次転記期待値。
- `expected/anomaly-expected.csv`: 異常系ケースの期待処理。

## チェック済みWorkbook

チェック済み `.xlsx` サンプルは `checked/` 配下に置きます。

- 日次サンプルは `checked/daily/<YYMM>/` 配下に置きます。
- 月次レイアウトサンプルは `checked/monthly/` 配下に置きます。
- `.xlsx` はCSV正本から生成される再生成可能な成果物です。
- `.xlsm` はVBA実装フェーズまで意図的に延期します。

## 再生成

期待結果CSVとチェック済み `.xlsx` サンプルを再生成する場合は、Codex workspaceのbundled Node runtimeで `tools/build-samples.mjs` を実行します。

通常実行では期待結果CSVを再生成し、不足しているチェック済みworkbookだけを作成します。既存のチェック済みworkbookを置き換える場合だけ `--force` を使います。

artifact-tool renderが利用できる環境で視覚確認を行う場合は `--render` を使います。
