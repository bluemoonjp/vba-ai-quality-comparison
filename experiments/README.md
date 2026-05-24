# 検証の共通規約

`experiments/` は、VBA品質比較の検証単位を置く場所です。表示上は「検証」と呼びます。

## 標準構成

各検証は `experiments/<番号>-<slug>/` に置きます。

- `README.md`: 人間向けの検証概要。
- `samples/`: 全条件で共有する合成データ、サンプルworkbook、期待結果。
- `generation-materials/`: モデルへ渡す共通仕様、prompt、制約、条件定義。条件別の生成物は置かない。
- `prompt-records/`: prompt投入記録、確認質問、回答、修正依頼、生成前に見えていたファイル範囲。
- `model-artifacts/`: 条件別の生成物、出力ログ、exported VBA。モデル成果物として `.xlsm` は置かない。
- `verification/`: 静的レビュー、Excel実行、評価表、公開安全点検。
- `reports/`: 日本人向けの比較レポート。生成用素材や条件別生成物とは混ぜない。

生成用素材、中間生成物、人間向けレポートは分けます。モデルに渡す入力は `generation-materials/` と `samples/` に集約し、投入履歴は `prompt-records/`、モデルから得た成果物は `model-artifacts/`、人間が比較して書く結論は `reports/` に置きます。

## 実験統制

- 全条件に同じ `generation-materials/` と `samples/` を使う。
- 生成前に見えるファイル範囲を記録する。
- 他条件の成果物、レビュー、評価結果を見せない。
- 確認質問は制限しないが、質問と回答をすべて記録する。
- 修正依頼は1回までにする。
- first outputからfinal outputまでの差分理由を記録する。
- 評価前に人間がモデル成果物を修正しない。修正した場合は比較対象から外すか、汚染として記録する。
- 先行条件の情報が見えてしまった場合は、`prompt-records/` または `verification/` に汚染内容と採用判断を残す。

## 記録と保存

prompt投入時は、条件ID、実行日、モデル/推論設定、投入prompt、参照したファイル、確認質問、回答、修正依頼、first output保存先、final output保存先、実験汚染の有無を記録します。

モデル成果物はMarkdownと `.bas` / `.cls` などのテキスト形式を正本にします。検証用 `.xlsm` や一時workbookはscratch扱いにし、公開成果物として保存しません。

## 役割

- Orchestrator: Issue整理、依頼、受領、検収、レビュー依頼。
- Explorer: 事実確認、既存ファイル調査、検証前の読み取り専用確認。
- Worker: 仕様、サンプル、prompt、生成物、評価表、レポートの作成。
- Reviewer: high reasoningで成果物と実験統制をレビューする。

Orchestratorは実ファイル作成、VBA生成、評価本文作成を直接担当しません。Orchestratorが担当するのは、作業依頼、進捗整理、成果物の受領、公開境界と実験統制の検収、Reviewer highへの確認依頼、Issue整理です。

## 公開境界

実データ、実パス、顧客名、社員名、secret、private source materialは扱いません。workbookを保存する場合は `docs/public-boundary.md` の公開前チェックを通過させます。
