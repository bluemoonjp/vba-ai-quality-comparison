# 検証の共通規約

`experiments/` は、VBA品質比較の検証単位を置く場所です。表示上は「検証」と呼びます。

## 標準構成

各検証は `experiments/<番号>-<slug>/` に置きます。

- `README.md`: 人間向けの検証概要。
- `samples/`: 合成データ、サンプルworkbook、期待結果。
- `generation-materials/`: モデルへ渡す仕様書、prompt、条件定義。
- `prompt-records/`: prompt投入記録、確認質問、回答、修正依頼。
- `model-artifacts/`: 条件別の生成物、出力ログ、exported VBA。
- `verification/`: 静的レビュー、Excel実行、評価表。
- `reports/`: 日本人向けの比較レポート。

## 実験統制

- 全条件に同じ `generation-materials/` と `samples/` を使う。
- 生成前に見えるファイル範囲を記録する。
- 他条件の成果物、レビュー、評価結果を見せない。
- 確認質問は制限しないが、質問と回答をすべて記録する。
- 修正依頼は1回までにする。
- first outputからfinal outputまでの差分理由を記録する。

## 役割

- Orchestrator: Issue整理、依頼、受領、検収、レビュー依頼。
- Explorer: 事実確認、既存ファイル調査、検証前の読み取り専用確認。
- Worker: 仕様、サンプル、prompt、生成物、評価表、レポートの作成。
- Reviewer: high reasoningで成果物と実験統制をレビューする。

Orchestratorは実ファイル作成、VBA生成、評価本文作成を直接担当しません。

## 公開境界

実データ、実パス、顧客名、社員名、secret、private source materialは扱いません。workbookを保存する場合は `docs/public-boundary.md` の公開前チェックを通過させます。
