# Issue #5 生成ログ収集ガイド

## 使うPrompt

各AI条件へ送るpromptは `prompts/common-task-prompt.md` の全文だけです。実行時点の正本commitは `78fbc6a` です。

## 実行順

1. `chatgpt-instant`
2. `chatgpt-deep`
3. `codex-medium`
4. `codex-xhigh`
5. `codex-multi-agent-review`

各条件は新規セッションまたは新規画面で実行します。先行条件の出力、修正依頼、評価メモを後続条件に見せません。

## 収集手順

1. 対象AIに `prompts/common-task-prompt.md` をそのまま貼ります。
2. 仕様確認質問が出たら、このスレッドに質問文を貼ります。
3. このスレッドで作った標準回答を対象AIへ貼ります。
4. 初回出力をこのスレッドに全文貼ります。
5. このスレッドで静的確認を行い、1回だけの修正依頼を作ります。
6. 修正依頼を対象AIへ貼ります。
7. 最終出力をこのスレッドに全文貼ります。
8. `final-output.md` と `final-vba.bas` に未編集で保存します。

## 保存ルール

- AI出力は人手修正しません。
- `final-vba.bas` は最終出力内のVBAコードブロックを抽出するだけです。
- 抽出時にコード内容は変えません。
- 追加説明、ログ、検証メモはMarkdownに分けて保存します。
- Excel実行検証と採点はIssue #6以降で行います。

## 未収集の扱い

このscaffold時点では全条件が未収集です。実際の出力を受け取ったら、各条件ディレクトリ内の `未収集` placeholderを実データに置き換えます。
