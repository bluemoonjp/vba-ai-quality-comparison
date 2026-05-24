# 条件定義

この検証では、同じ仕様、同じサンプル、同じ期待結果を使って、Codexの使い方による品質差を比較します。

## 共通ルール

- 確認質問は回数制限なし。
- 修正依頼は1回まで。
- 他条件の成果物、レビュー、評価結果を見せない。
- 生成前に見えていたファイル範囲を記録する。
- モデル成果物はMarkdownと `.bas` で保存する。
- `.xlsm` はモデル成果物として保存しない。

## 条件

| condition_id | 目的 | 設定 |
| --- | --- | --- |
| `codex-chatgpt-medium` | ChatGPT代替としてのCodex利用 | reasoning medium |
| `codex-chatgpt-xhigh` | ChatGPT代替としての高推論Codex利用 | reasoning xhigh |
| `codex-plan-medium` | 単純な計画から実装まで、ファイル直接操作を許可 | reasoning medium |
| `codex-plan-xhigh` | 単純な計画から実装まで、ファイル直接操作を許可 | reasoning xhigh |
| `codex-multi-review` | マルチエージェントでレビューしながら実装 | Orchestrator high, Explorer medium, Worker medium, Reviewer high |

## 記録すること

- 実行日。
- モデル/推論設定。
- 投入prompt。
- 参照可能だったファイル。
- 確認質問と回答。
- first output。
- 修正依頼。
- final output。
- `final-vba.bas`。
- 汚染防止チェック。
