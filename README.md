# vba-ai-quality-comparison

AIの使い方によって、VBA成果物の品質がどのように変わるかを比較する公開実験用リポジトリです。

このリポジトリはPublic公開を前提にしています。扱う支店名、業務名、フォルダ構成、クレーム件数はすべて架空データです。実データ、個人情報、社内パス、secret、実ブック由来のスクリーンショットは追加しません。

## 目的

同じVBA課題を、次の条件で作成させ、成果物の品質差を比較します。

- ChatGPT Instant
- ChatGPT Thinking extended
- Codex、単純な計画から実装、reasoning medium
- Codex、単純な計画から実装、reasoning xhigh
- Codex、マルチエージェントレビュー構成
  - Orchestrator: high
  - Explorer: medium
  - Reviewer: high

課題は、月次ブック `月次クレーム集計YYMM.xlsm` に埋め込むマクロです。マクロはサブフォルダ内の日次ブック `クレーム集計YYMMDD.xlsx` を読み、支店別・業務別のクレーム件数を月次表へ転記します。

## 現在の状態

初期計画、公開境界、実験プロトコル、合成サンプル、AI向け課題仕様、日本語化を整備済みです。以降の作業はGitHub Issuesで管理します。

## 実験プロトコル

- 全条件で同じ課題仕様、サンプルデータ、期待結果、評価基準を使います。
- 各条件で許可するやり取りは、仕様確認1回と生成後の修正依頼1回までです。
- 評価前の人手によるコード修正は禁止です。保存形式変換やVBAのexportは、ログに記録した場合のみ許可します。
- モデル、モード、実行日、prompt、確認質問、回答、修正依頼、最終出力、検証結果を記録します。

## リポジトリ構成

- `docs/`: 実験設計、課題仕様、評価基準、結果まとめ。
- `prompts/`: 共通prompt、条件別promptログ、確認質問への標準回答。
- `outputs/`: 条件別の生成ログと最終成果物。
- `samples/`: 合成サンプル、CSV正本、期待結果、チェック済みExcelブック。
- `src/`: exportしたVBAモジュール。
- `tests/`: Excel上の検証観点と評価表。
- `.github/issues/`: 初期Issue本文の控え。

## 公開境界

prompt、ログ、workbook、exported VBA、スクリーンショット、生成物を追加する前に、[公開境界](docs/public-boundary.md)を確認してください。Issue #1 のscaffold点検は [docs/scaffold-audit-issue-1.md](docs/scaffold-audit-issue-1.md) に、日本語化の点検は [docs/localization-audit-issue-8.md](docs/localization-audit-issue-8.md) に記録しています。

## Git境界

このディレクトリは、親 `_Workspace` フォルダ配下にある独立Gitリポジトリです。親workspaceのcommitやIssueと混ぜないでください。

## 公開前チェック

成果物をpushまたは公開する前に、次を確認します。

- 実在の支店、顧客、社員、会社、クレーム内容が含まれていない。
- 実パス、ユーザー名、OneDriveパス、端末名が含まれていない。
- workbookの隠しシート、ドキュメントプロパティ、コメント、メタデータに非公開情報がない。
- secret、token、APIレスポンス、private source materialが含まれていない。
- VBAはレビューしやすいテキスト形式でexportする。`.xlsm` は検証サンプルとして慎重に扱う。

## ライセンス

MIT Licenseです。詳細は `LICENSE` を参照してください。
