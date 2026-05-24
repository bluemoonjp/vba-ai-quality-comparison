# Issueテンプレート

## 目的

## 対象範囲

## 対象外

## Owner role

通常は Orchestrator。Orchestratorは依頼、受領、検収、レビュー依頼、Issue整理だけを担当する。

## Executor role

実作業を行う role を明記する。生成Issueでは該当条件のWorker/モデルを指定し、Orchestratorは `final-vba.bas` を作らない。

## Reviewer role

重要成果物は Reviewer high が確認する。

## 成果物パス

生成用素材、サンプル、prompt記録、モデル成果物、検証記録、人間向けレポートを混ぜずに書く。

- 生成用素材:
- サンプル:
- prompt記録:
- モデル成果物:
- 検証記録:
- 人間向けレポート:

## 依存Issue

## 公開リスク

実データ、実パス、secret、private source material、未確認Officeバイナリの混入可能性を書く。workbookを保存する場合は `docs/public-boundary.md` の確認方法も書く。

## 実験汚染防止

生成前に見えるファイル範囲、見せない成果物、条件間で共有してよい入力、汚染が起きた場合の記録先を書く。

## 次アクション

## 完了条件

- Orchestratorが実作業を直接担当していないことを確認する。
- 成果物パスが、生成用素材、サンプル、prompt記録、モデル成果物、検証記録、人間向けレポートに分かれている。
- prompt、確認質問、回答、修正依頼、参照可能ファイル、汚染防止チェックの記録先が分かる。
- 公開境界の確認結果または確認予定が残っている。
