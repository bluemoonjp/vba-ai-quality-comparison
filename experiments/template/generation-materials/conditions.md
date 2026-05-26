# 条件定義

このファイルは、各条件の実行条件と比較上の前提を定義します。

## 標準5条件

002 以降も初回比較と同じ条件で比較する場合は、次の5条件を初期値として使います。

| condition_id | 目的 | 記録先 |
| --- | --- | --- |
| `codex-chatgpt-medium` | 通常チャット相当、medium reasoning | `prompt-records/codex-chatgpt-medium.md` |
| `codex-chatgpt-xhigh` | 通常チャット相当、xhigh reasoning | `prompt-records/codex-chatgpt-xhigh.md` |
| `codex-plan-medium` | 事前計画あり、medium reasoning | `prompt-records/codex-plan-medium.md` |
| `codex-plan-xhigh` | 事前計画あり、xhigh reasoning | `prompt-records/codex-plan-xhigh.md` |
| `codex-multi-review` | Worker / Reviewer を使うマルチエージェント構成 | `prompt-records/codex-multi-review.md` |

## 任意条件へ差し替える場合

- `condition_id` は小文字英数字とハイフンを使い、ディレクトリ名とファイル名にそのまま使える形にする。
- 条件を増減した場合は、`prompt-records/`, `model-artifacts/`, `verification/results-matrix.md`, `reports/human-report.md` の行も揃える。
- 条件差は、モデル、推論設定、prompt作成方法、レビュー有無など、比較したい要素だけに絞る。
- すべての条件で同じ `samples/` と `generation-materials/task-spec.md` を使う。
- 先行条件の成果物、レビュー、評価結果を後続条件に見せない。

## 共通ルール

- 確認質問は制限しないが、質問と回答をすべて記録する。
- 最終出力後の修正依頼は1回までにする。
- 評価前に人間が生成 VBA を修正した場合は、比較対象から外すか、汚染として記録する。
- 生成物の正本は Markdown と `.bas` / `.cls` などのテキスト形式にする。
