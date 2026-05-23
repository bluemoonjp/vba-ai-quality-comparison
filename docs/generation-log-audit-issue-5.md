# Issue #5 生成ログ収集Audit

## 目的

Issue #5では、5つのAI条件から生成ログと最終成果物を同じ形式で収集します。

## 今回固定した運用

- 条件別Issueには分割せず、Issue #5で一括管理します。
- 各AI条件は独立した新規セッションで実行します。
- このスレッドではVBAを生成せず、ログ整理、標準回答作成、修正依頼作成、保存、closeoutだけを担当します。
- 全条件で `prompts/common-task-prompt.md` の全文だけを初回promptとして使います。
- 仕様確認質問は1回まで、修正依頼は1回までです。

## 追加した受け皿

- `outputs/collection-guide-issue-5.md`
- `outputs/chatgpt-instant/`
- `outputs/chatgpt-deep/`
- `outputs/codex-medium/`
- `outputs/codex-xhigh/`
- `outputs/codex-multi-agent-review/`

各条件ディレクトリには `run-log.md`, `first-output.md`, `verification-before-correction.md`, `correction-request.md`, `final-output.md`, `final-vba.bas` を置きます。

## 現在の収集状態

全条件とも未収集です。AI出力を受け取り次第、placeholderを未編集の出力へ置き換えます。

## 後続作業

- ユーザーが各AI条件に共通promptを貼り、出力をこのスレッドへ貼ります。
- 仕様確認質問が出た場合は、`prompts/clarification-answer-bank.md` の範囲で回答します。
- 初回出力を保存後、同じ静的観点で確認し、1回だけの修正依頼を作ります。
- 最終出力と抽出VBAを保存した後、全条件が揃った段階でIssue #5をcloseします。
