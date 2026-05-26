# prompt記録

条件ごとの prompt 投入記録を保存します。

## 標準5条件のファイル名

- `codex-chatgpt-medium.md`
- `codex-chatgpt-xhigh.md`
- `codex-plan-medium.md`
- `codex-plan-xhigh.md`
- `codex-multi-review.md`

任意条件へ差し替える場合は、`generation-materials/conditions.md` の `condition_id` と同じファイル名にします。

## 必須項目

- 条件ID。
- 実行日。
- モデル/推論設定。
- 実行形態。
- 投入prompt。
- 参照したファイル。
- 生成前に参照可能だったファイル範囲。
- 見せなかったファイル範囲。
- 確認質問。
- 回答。
- 修正依頼。
- first output保存先。
- final output保存先。
- `final-vba.bas` 保存先。
- 実験汚染の有無。

## 手動記録テンプレート

````markdown
# <condition_id> prompt記録

## 実行情報

- 条件ID:
- 実行日:
- モデル/推論設定:
- 実行形態:

## 参照可能ファイル

- generation-materials:
- samples:
- 見せなかったファイル:

## 投入prompt

```text

```

## 確認質問と回答

- 質問数:
- 質問:
- 回答:
- 標準回答を使った場合の理由:

## first output

- 保存先:
- 要約:

## 修正依頼

- 依頼内容:
- 依頼回数: 1回まで
- 修正依頼なしの場合の理由:

## final output

- Markdown保存先:
- `final-vba.bas` 保存先:
- 入口プロシージャの明示:

## 汚染防止チェック

- 他条件の成果物、レビュー、評価結果を見ていない:
- 追加指示の有無:
- 見えてはいけない情報が見えた場合の内容:
- 採用判断:
````

確認質問が0回、修正依頼が0回だった場合も、比較不能な観点として後で読めるように明記します。
