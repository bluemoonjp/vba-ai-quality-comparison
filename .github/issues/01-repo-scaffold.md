## 目的

Public repositoryとして安全に実験を進めるための初期構成と公開境界を作る。

## 背景

このrepoはAI使用方法によるVBA成果物品質差を比較する公開実験である。実データや個人情報を含めず、独立repoとして親 `_Workspace` と混ぜない。

## 対象範囲

- README、LICENSE、`.gitignore`、AGENTS、docs/samples/src/outputs/tests構成
- Public公開前チェック
- 独立Git repoとしての境界説明

## 対象外

- 実験本体の実行
- VBA生成
- Excelサンプル作成

## 方針または作業案

初期scaffoldを確認し、不足する公開安全ルールと再現手順入口を補う。

## 確認が必要な操作

GitHub repo作成、初回push、Issue作成はユーザー確認済みの範囲で実施する。

## リスク

実パス、実データ、Excelメタデータ、不要なバイナリがPublic repoに混入する可能性。

## 次アクション

初期構成と公開安全チェックをレビューし、必要な修正を行う。

## 完了条件

- READMEから目的、実験条件、Git boundary、公開安全線が分かる。
- MIT Licenseがある。
- 実データ、個人情報、社内情報、secretが含まれていない。
- 親 `_Workspace` repoと独立repoの差分が混ざっていない。

## 未確認

- `.xlsm`を初期repoに含めるかは後続Issueで判断する。
