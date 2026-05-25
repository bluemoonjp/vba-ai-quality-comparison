# codex-multi-review prompt record

- 実行日: 2026-05-25
- 条件ID: `codex-multi-review`
- モデル/推論設定: Orchestrator high / Explorer medium / Worker medium / Reviewer high
- 実行形態: Orchestratorが依頼と統合、Explorerが許可資料を調査、Workerが成果物作成、Reviewer highがレビュー完了し指摘なしと判断

## 投入prompt

```text
あなたはExcel VBAに詳しい開発者です。

このpromptは全条件で同一に使います。他条件の生成物、レビュー、評価結果は参照しないでください。

添付または参照できるサンプルファイルと仕様書を確認し、月次クレーム集計表へ日次クレーム件数を転記するVBAを作成してください。

入力:
- 仕様書: generation-materials/task-spec.md
- 制約: generation-materials/prompt-constraints.md
- 条件定義: generation-materials/conditions.md
- サンプル: samples/workbooks/
- 期待結果: samples/expected/monthly-expected.csv

参照してよいファイルは、このpromptで示した入力と、必要に応じて samples/README.md までです。他条件の model-artifacts/, prompt-records/, verification/, reports/ は参照しないでください。

出力:
- 標準モジュールへ貼り付けられる完全なVBAコード。
- .bas として保存できる形のコード。
- 入口プロシージャ名。
- 実行方法。
- 参照設定が必要かどうか。
- ログ出力方法。
- 再実行しても二重加算しない理由。
- Markdownログとして残せる説明。

注意:
MsgBoxで処理を止める実装、実ローカルパスの固定、外部通信、サンプルworkbookの構造変更を前提にした実装は禁止です。

分からないことは確認質問してください。確認質問の回数は制限しません。ただし、最終出力後の修正依頼は1回だけです。
```

## 参照可能ファイル

- `AGENTS.md`
- `README.md`
- `experiments/001-claims-monthly-simple/generation-materials/task-spec.md`
- `experiments/001-claims-monthly-simple/generation-materials/prompt-template.md`
- `experiments/001-claims-monthly-simple/generation-materials/prompt-constraints.md`
- `experiments/001-claims-monthly-simple/generation-materials/conditions.md`
- `experiments/001-claims-monthly-simple/generation-materials/clarification-answer-bank.md`
- `experiments/001-claims-monthly-simple/samples/README.md`
- `experiments/001-claims-monthly-simple/samples/source/`
- `experiments/001-claims-monthly-simple/samples/expected/`
- `experiments/001-claims-monthly-simple/samples/workbooks/`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/README.md`

## 参照したファイル

- `AGENTS.md`
- `README.md`
- `experiments/001-claims-monthly-simple/generation-materials/task-spec.md`
- `experiments/001-claims-monthly-simple/generation-materials/prompt-template.md`
- `experiments/001-claims-monthly-simple/generation-materials/prompt-constraints.md`
- `experiments/001-claims-monthly-simple/generation-materials/conditions.md`
- `experiments/001-claims-monthly-simple/generation-materials/clarification-answer-bank.md`
- `experiments/001-claims-monthly-simple/samples/README.md`
- `experiments/001-claims-monthly-simple/samples/source/branches.csv`
- `experiments/001-claims-monthly-simple/samples/source/business-lines.csv`
- `experiments/001-claims-monthly-simple/samples/source/daily-records.csv`
- `experiments/001-claims-monthly-simple/samples/expected/monthly-expected.csv`
- `experiments/001-claims-monthly-simple/samples/workbooks/`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/README.md`

## 確認質問と回答

確認質問は0回。

標準回答として利用可能だった内容:
- 月次ファイルと同じフォルダ配下に `daily/2605/` がある。
- 初回検証の対象月は `2605`。
- MsgBoxで止めず、ログシートまたはログ表に記録する。
- 追加参照は避け、標準VBAまたはlate bindingを優先する。
- 最終出力後の修正依頼は1回だけ。

## 修正依頼

修正依頼0回。

## 成果物パス

- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/orchestrator-log.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/explorer-log.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/worker-log.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/reviewer-log.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/first-output.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/correction-request.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/final-output.md`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/final-vba.bas`
- `experiments/001-claims-monthly-simple/model-artifacts/codex-multi-review/notes.md`

## role構成

- Orchestrator high: 依頼、受領、統合判断。VBA本文の改善・補完はしない。
- Explorer medium: 許可資料の調査と公開可能な仕様整理。
- Worker medium: VBA生成、first/final output、final-vba.bas抽出、role logs、prompt record、notes保存。
- Reviewer high: レビュー完了。指摘なし。

## 汚染防止チェック

- 生成前に参照可能だったファイル範囲を記録した。
- 他条件の `prompt-records/`, `model-artifacts/`, `verification/`, `reports/` は参照していない。
- #15から#18の成果物、レビュー、生成済みVBA本文は参照していない。
- 実パス、ユーザー名、同期サービス名、機密情報、認証情報、非公開情報は成果物に含めていない。
- `.xlsm` は作成していない。
