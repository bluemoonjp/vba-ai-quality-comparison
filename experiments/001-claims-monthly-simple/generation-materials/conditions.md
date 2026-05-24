# 条件定義

この検証では、同じ仕様、同じサンプル、同じ期待結果を使って、Codexの使い方による品質差を比較します。

## 共通ルール

- 確認質問は回数制限なし。
- 修正依頼は1回まで。
- 他条件の成果物、レビュー、評価結果を見せない。
- 生成前に見えていたファイル範囲を記録する。
- モデル成果物はMarkdownと `.bas` で保存する。
- `.xlsm` はモデル成果物として保存しない。

## 許可ファイル範囲

全条件で共通して参照してよいファイルは次の範囲に限ります。

- `generation-materials/task-spec.md`
- `generation-materials/prompt-template.md`
- `generation-materials/prompt-constraints.md`
- `generation-materials/conditions.md`
- `generation-materials/clarification-answer-bank.md`
- `samples/README.md`
- `samples/source/`
- `samples/expected/`
- `samples/workbooks/`

条件別の `prompt-records/`, `model-artifacts/`, `verification/`, `reports/` は、該当条件の生成前には見せません。見えてしまった場合は汚染として記録します。

## 条件

| condition_id | 目的 | モデル/推論設定 | 実行形態 |
| --- | --- | --- | --- |
| `codex-chatgpt-medium` | ChatGPT代替としてのCodex利用 | reasoning medium | prompt応答のみ。ファイル直接編集はしない。 |
| `codex-chatgpt-xhigh` | ChatGPT代替としての高推論Codex利用 | reasoning xhigh | prompt応答のみ。ファイル直接編集はしない。 |
| `codex-plan-medium` | 単純な計画から実装まで、ファイル直接操作を許可 | reasoning medium | 計画後にファイル操作を許可。 |
| `codex-plan-xhigh` | 単純な計画から実装まで、ファイル直接操作を許可 | reasoning xhigh | 計画後にファイル操作を許可。 |
| `codex-multi-review` | マルチエージェントでレビューしながら実装 | Orchestrator high, Explorer medium, Worker medium, Reviewer high | Orchestratorが依頼と統合、Workerが実作業、Reviewerが確認。 |

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

## 汚染防止チェック

各条件のprompt記録に、次を記録します。

- 生成前に参照可能だったファイル範囲。
- 他条件の `prompt-records/`, `model-artifacts/`, `verification/`, `reports/` を見ていないこと。
- 追加指示、補足説明、修正依頼の有無。
- 見えてはいけない情報が見えた場合の内容と採用判断。
