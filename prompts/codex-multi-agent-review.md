# Codex Multi-Agent Review Prompt

## 実行メタデータ

- condition id: `codex-multi-agent-review`
- オーケストレーター: high
- explorer: medium
- reviewer: high
- 実行日: 2026-05-24
- 実行者: user

## 初回Prompt

この条件では、以下のマルチエージェント実行方式指定を先頭に置き、その後に `prompts/common-task-prompt.md` の全文を続けて送信する。

```text
あなたは Codex multi-agent 条件のオーケストレーターです。

実験条件:
- オーケストレーター: reasoning high
- explorer: reasoning medium
- reviewer: reasoning high
- オーケストレーターは、explorer と reviewer の提案を統合する役割です。
- explorer は仕様の読み取り、設計方針、VBA実装上の注意点、落とし穴を洗い出してください。
- reviewer は explorer案と最終コード案をレビューし、仕様漏れ、再実行安全性、Excel/VBA慣習、エラー処理、保守性の観点で指摘してください。
- 改善ループは1回だけ行ってください。
- 最終回答には、標準モジュールに貼り付けられる完全なVBAコードと、日本語の説明を含めてください。
- 先行条件の出力や修正依頼は参照しないでください。
- 仕様確認は1回まで、生成後の修正依頼も1回までです。

以下の共通課題Promptに従ってください。

---
```

## 共通Prompt

上記に続けて、`prompts/common-task-prompt.md` の全文を貼る。

## 仕様確認質問

未使用。

## 仕様確認回答

未使用。

## Explorer出力

初回出力では、Explorer出力は個別ブロックとしては分離されず、最終回答に統合された。

## Reviewer出力

初回出力では、Reviewer出力は個別ブロックとしては分離されず、最終回答に統合された。

## 初回統合出力

`outputs/codex-multi-agent-review/first-output.md` に保存。

## 修正依頼

`outputs/codex-multi-agent-review/correction-request.md` に保存。

## 最終出力

`outputs/codex-multi-agent-review/final-output.md` に保存。
