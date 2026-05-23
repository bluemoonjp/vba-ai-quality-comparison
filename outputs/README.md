# 出力

AI条件ごとの生成結果をここへ保存します。

各condition folderには、次を同じ形式で残します。

- `run-log.md`
- export済みVBAファイル
- 初回出力
- 修正依頼
- 最終出力
- 検証メモ

各条件の記録は `run-log-template.md` を起点にします。

private screenshot、実workbook、実パス、未確認ログは保存しません。

## Issue #5 収集状態

Issue #5 の収集手順は `collection-guide-issue-5.md` を参照します。

現在の受け皿:

- `chatgpt-instant/`
- `chatgpt-deep/`
- `codex-medium/`
- `codex-xhigh/`
- `codex-multi-agent-review/`

scaffold時点では各条件のAI出力は未収集です。実出力を受け取ったら、各条件ディレクトリ内のplaceholderを未編集の出力へ置き換えます。
