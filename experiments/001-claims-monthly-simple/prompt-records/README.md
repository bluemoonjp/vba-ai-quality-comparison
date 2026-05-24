# prompt記録

各条件のprompt投入記録をここに保存します。

## ファイル名

- `codex-chatgpt-medium.md`
- `codex-chatgpt-xhigh.md`
- `codex-plan-medium.md`
- `codex-plan-xhigh.md`
- `codex-multi-review.md`

## 必須項目

- 条件ID。
- 実行日。
- モデル/推論設定。
- 投入prompt。
- 参照したファイル。
- 生成前に参照可能だったファイル範囲。
- 確認質問。
- 回答。
- 修正依頼。
- first output保存先。
- final output保存先。
- `final-vba.bas` 保存先。
- 実験汚染の有無。

自動記録できない場合は、同じ項目を手動で記録します。

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

- 質問:
- 回答:

## first output

- 保存先:

## 修正依頼

- 依頼内容:
- 依頼回数: 1回まで

## final output

- Markdown保存先:
- `final-vba.bas` 保存先:

## 汚染防止チェック

- 他条件の成果物、レビュー、評価結果を見ていない:
- 追加指示の有無:
- 見えてはいけない情報が見えた場合の内容:
- 採用判断:
````

確認質問は制限しません。最終出力後の修正依頼は1回だけです。
