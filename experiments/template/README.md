# 反復検証テンプレート

このディレクトリは、`experiments/<NNN>-<slug>/` を作るためのコピー元です。

## 使い方

1. このディレクトリ全体を `experiments/<NNN>-<slug>/` にコピーする。
2. `<NNN>`, `<slug>`, `<experiment_title>`, `<condition_id>` などのプレースホルダを置き換える。
3. `samples/` に合成データ、期待結果、公開前チェック済みの sample workbook を置く。
4. `generation-materials/` に全条件で共通に見せる仕様、prompt、制約、条件定義を置く。
5. 条件ごとの投入履歴を `prompt-records/` に記録する。
6. 条件ごとの生成物を `model-artifacts/<condition_id>/` に保存する。
7. 評価結果を `verification/` に記録し、人間向け比較レポートを `reports/` に書く。

## コピー後に必ず置き換える項目

- `<NNN>-<slug>`: 検証ID。
- `<experiment_title>`: 検証名。
- `<task_summary>`: VBA課題の短い説明。
- `<condition_id>`: 条件ID。標準5条件を使わない場合は `generation-materials/conditions.md` と各記録先を揃える。
- `<sample_file>` / `<expected_file>`: サンプルと期待値のファイル名。
- `<entry_procedure>`: 実行する入口プロシージャ。未確定の場合は評価時に判定する。
- `<task_specific_check>`: 課題固有の評価軸。月次集計なら `monthly_total` のように実験開始前に固定する。

## 公開境界チェック

- 実データ、実パス、顧客名、社員名、secret、private source material を含めない。
- ローカル絶対パス、ユーザー名、端末名、スクリーンショットを公開成果物に残さない。
- `.xlsm` や検証用 scratch workbook は成果物にしない。
- sample workbook を保存する場合は `docs/public-boundary.md` の workbook 公開前チェックを通過させる。
- 初回検証など過去実験の成果物、ログ、VBA、workbook はコピーしない。再利用するのは構造と手順だけにする。

## 標準の成果物境界

- `samples/`: 全条件で共有する合成データ、sample workbook、期待値。
- `generation-materials/`: モデルへ渡す共通仕様、prompt、制約、標準回答、条件定義。
- `prompt-records/`: 条件別の投入記録、確認質問、回答、修正依頼。
- `model-artifacts/`: 条件別の生成物、出力ログ、exported VBA。
- `verification/`: 評価手順、実行ログ、評価表、公開安全点検。
- `reports/`: 公開向け比較レポート。
