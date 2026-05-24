# vba-ai-quality-comparison

AIの使い方によって、VBA成果物の品質がどのように変わるかを比較する公開実験用リポジトリです。

このリポジトリはPublic公開を前提にしています。扱う支店名、業務名、フォルダ構成、クレーム件数はすべて架空データです。実データ、個人情報、社内パス、secret、実ブック由来のスクリーンショットは追加しません。

## 目的

同じVBA課題を複数のCodex利用方法で作成させ、成果物の品質差を比較します。比較対象は、モデルそのものだけではなく、計画、ファイル操作、レビュー、マルチエージェント構成の使い方です。

初回検証は `experiments/001-claims-monthly-simple/` です。日次ブック `クレーム集計YYMMDD.xlsx` から月次集計表 `月次クレーム集計YYMM.xlsm` へ、支店+業務別のクレーム件数を転記するVBAを作成させます。

## 現在の状態

旧計画は破棄し、`experiments/` 配下で検証単位ごとに管理する構成へ再構築中です。作業はGitHub Issuesで管理し、Issueを順番に処理すれば初回検証とレポート作成まで進められる形にします。

## GitHub Issues

新しい作業トラッカーは Issue #9 以降です。

- #9: 旧計画の破棄対象を確定する。
- #10: 旧Issue Closeとリポジトリ再構築を実行する。
- #11: 新実験プロトコルとフォルダ規約を定義する。
- #12: 初回検証サンプルを作成する。
- #13: モデル投入用仕様書と共通promptを作成する。
- #14: 条件定義とprompt記録ルールを整備する。
- #15-#19: 5条件のモデル生成を実施する。
- #20: 初回検証を実Excelで評価する。
- #21: 初回比較レポートを作成する。
- #22: 反復検証テンプレートを整備する。

## リポジトリ構成

- `docs/public-boundary.md`: 公開境界とworkbook保存ルール。
- `experiments/README.md`: 反復検証の共通規約。
- `experiments/001-claims-monthly-simple/`: 初回検証。
- `experiments/001-claims-monthly-simple/generation-materials/`: モデル投入用仕様書、prompt、条件定義。
- `experiments/001-claims-monthly-simple/prompt-records/`: prompt投入記録。
- `experiments/001-claims-monthly-simple/model-artifacts/`: 条件別の生成物と `.bas`。
- `experiments/001-claims-monthly-simple/verification/`: 検証計画と評価表。
- `experiments/001-claims-monthly-simple/reports/`: 人間向けレポート。

## 実験プロトコル

- 全条件で同じ課題仕様、サンプルデータ、期待結果、評価基準を使います。
- 確認質問の回数は制限しません。修正依頼は各条件1回までです。
- 評価前の人手によるコード修正は禁止します。
- モデル、モード、実行日、prompt、確認質問、回答、修正依頼、最終出力、検証結果を記録します。
- 先行条件の成果物やレビュー結果を後続条件に見せません。

## エージェント運用

Orchestratorは作業実体を持たず、依頼、進捗整理、受領、検収、レビュー依頼、Issue整理を担当します。仕様作成、ファイル作成、VBA生成、評価本文作成はWorker / Explorerが担当し、重要成果物はReviewer highで確認します。

## 公開境界

prompt、ログ、workbook、exported VBA、スクリーンショット、生成物を追加する前に、[公開境界](docs/public-boundary.md)を確認してください。

モデル成果物として `.xlsm` は保存しません。サンプル workbook は公開チェック済みのものだけ `experiments/**/samples/workbooks/` に保存します。モデルが作った成果物は `.bas` とMarkdownログを正本にします。

## Git境界

このディレクトリは、親 `_Workspace` フォルダ配下にある独立Gitリポジトリです。親workspaceのcommitやIssueと混ぜないでください。
