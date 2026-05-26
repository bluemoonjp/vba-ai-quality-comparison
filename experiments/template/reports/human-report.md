# <experiment_title> 比較レポート

## 結論

<summary>

結論は、`verification/results-matrix.md` と `verification/excel-run-log.md` で確認できる範囲に限定します。正常系だけで評価した場合は、その限界を明記します。

## 実験概要と評価方法

- 検証ID:
- 課題:
- 条件:
- サンプル:
- 評価方法:
- 追加評価の有無:

## 比較表

`verification/results-matrix.md` を正本として引用します。

| condition_id | import | run | transfer | task_specific_check | rerun_safe | no_msgbox | major_defect |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `codex-chatgpt-medium` |  |  |  |  |  |  |  |
| `codex-chatgpt-xhigh` |  |  |  |  |  |  |  |
| `codex-plan-medium` |  |  |  |  |  |  |  |
| `codex-plan-xhigh` |  |  |  |  |  |  |  |
| `codex-multi-review` |  |  |  |  |  |  |  |

`task_specific_check` は実験開始前に課題へ合わせて固定した評価軸へ置き換えます。月次集計なら `monthly_total` のように、評価表と同じ名称を使います。

## 条件別の強み・弱み

- 強み:
- 弱み:
- 根拠:

標準5条件を使う場合は、条件ごとに見出しを追加します。任意条件へ差し替える場合は、`generation-materials/conditions.md` の `condition_id` と同じ見出しにします。

## 失敗しなかった点 / 差が出た点

- 失敗しなかった点:
- 差が出た点:
- 差が測れなかった点:

正確性で差が出なかった場合は、生成過程、構造、レビュー容易性、説明の厚み、次回複雑化時の伸びしろなど、根拠がある観点だけを書く。

## 確認質問と修正依頼の観察

- 確認質問:
- 標準回答の有無:
- 修正依頼:
- 改善度比較:

確認質問が0回または標準回答のみだった場合は、質問力の差を測れなかったと明記します。修正依頼が0回だった場合は、修正後改善度は比較不能とします。

## 限界

- サンプル規模:
- 正常系/異常系の範囲:
- 評価できなかった観点:
- 手動確認に残るリスク:

## 次回複雑化案

- 未知キー、空白、非数値、重複。
- 複数月、月またぎ、可変日数。
- シート欠落、対象外ファイル、保護列。
- 意図的な修正依頼を含めた改善度比較。
- あいまい仕様による確認質問の質の比較。

## 根拠

- `verification/results-matrix.md`
- `verification/excel-run-log.md`
- `generation-materials/conditions.md`
- `prompt-records/<condition_id>.md`
- `model-artifacts/<condition_id>/final-output.md`
- `model-artifacts/<condition_id>/final-vba.bas`
- `model-artifacts/<condition_id>/notes.md`
