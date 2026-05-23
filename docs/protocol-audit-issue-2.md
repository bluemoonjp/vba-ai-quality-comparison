# Issue #2 プロトコル監査

監査日: 2026-05-23

## 対象範囲

Issue #2 では、実験プロトコル、ログ形式、評価方式を固定しました。AI条件の実行、Excelサンプル作成、VBA生成は行っていません。

## 固定した判断

- 実験では、次の5つの安定した `condition id` を使う。
  - `chatgpt-instant`
  - `chatgpt-thinking-extended`
  - `codex-medium`
  - `codex-xhigh`
  - `codex-multi-agent-review`
- 各条件で許可するやり取りは、仕様確認1回と修正依頼1回まで。
- 採点はカテゴリごとに `0-4` 点とし、根拠メモを必須にする。
- 総合点は補助情報であり、それだけで結論を出さない。
- 重大欠陥、未実行項目、修正効果、公平性/汚染メモを記録する。

## 追加したテンプレート

- `outputs/run-log-template.md`
- `tests/evaluation-sheet-template.md`

これらは、Issue #5 の生成ログと Issue #6 の検証/評価記録の標準テンプレートです。

## 後続作業

- 共通promptの最終化は Issue #4 に依存。
- 合成workbook/sample作成は Issue #3 に依存。
- 実際のAI条件実行は Issue #5 に依存。
- Windows版Excelでの検証は Issue #6 に依存。

## 完了条件

Issue #2 は、次の状態で完了としました。

- 追加判断なしでプロトコルを実行できる。
- 採点尺度とカテゴリ別根拠の記録ルールが文書化されている。
- run-logと評価表テンプレートが用意されている。
- AI実行、Excelサンプル、VBA実装をこのIssueで行っていない。
