# model-artifacts

条件別の生成物を保存します。

## 標準構成

各条件は `model-artifacts/<condition_id>/` に置きます。

- `first-output.md`: 最初の出力。
- `final-output.md`: 修正依頼後または最終採用した出力。
- `final-vba.bas`: 評価対象の exported VBA。
- `notes.md`: 判断、制約、未確認事項、補足。
- `correction-request.md`: 修正依頼を出した場合の依頼内容。
- `worker-log.md`, `reviewer-log.md`, `orchestrator-log.md`: マルチエージェント条件で必要な場合だけ保存する。

## 保存ルール

- 評価前に人間が `final-vba.bas` を修正しない。
- `.xlsm`、scratch workbook、スクリーンショットを保存しない。
- ローカル絶対パス、ユーザー名、端末名、secret をログに残さない。
- `.bas` が文字化け対策のために一時変換された場合でも、変換後 scratch copy は保存しない。
- 条件間で生成物を見せ合わない。

## 任意条件への対応

`generation-materials/conditions.md` の `condition_id` ごとに同じ構成を作ります。条件を増減した場合は、`prompt-records/` と `verification/results-matrix.md` も同じ条件数に揃えます。
