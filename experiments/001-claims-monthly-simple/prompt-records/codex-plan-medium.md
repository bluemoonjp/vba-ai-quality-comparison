# codex-plan-medium prompt record

## 基本情報

- Issue: #17
- 条件ID: `codex-plan-medium`
- モデル/推論設定として記録する条件: reasoning medium
- 実行形態: 単純な計画から実装まで。ファイル直接操作を許可。
- 実行日: 2026-05-25
- branch: `codex/issue-17-codex-plan-medium`
- 起点: Orchestrator指示により `origin/main` 起点として記録。
- #15/#16成果物: Orchestrator指示により、#17 branch 起点では #15/#16成果物は存在しなかったものとして記録。

## 投入prompt / 指示要約

共通promptテンプレートに従い、`generation-materials/task-spec.md`、`prompt-constraints.md`、`conditions.md`、サンプル、期待結果を参照して、月次クレーム集計表へ日次クレーム件数を転記するVBAを生成した。

追加のWorker向け指示として、短い計画を立ててから所有範囲の成果物ファイルを直接作成し、`first-output.md`、`correction-request.md`、`final-output.md`、`final-vba.bas`、`notes.md`、本prompt recordへ保存した。

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
- `experiments/001-claims-monthly-simple/samples/workbooks/` 配下のファイル一覧

## 確認質問と回答

- 確認質問: なし。
- 標準回答採用: あり。`clarification-answer-bank.md` の標準回答に従い、日次フォルダは月次ファイルと同じフォルダ配下の `daily/2605/`、対象月は `2605`、ログはログシート、追加参照は避ける方針を採用した。

## 計画から実装までのログ

1. 計画: 月次表の支店コード + 業務コードから行番号を引く辞書を作り、日次ファイルごとに同じキーで件数を合算して日別列へ転記する。
2. 判断: 転記前に `E:AI` と `AJ` をクリアし、月合計は転記後に再計算することで再実行安全性を満たす。
3. 判断: 追加参照を避けるため、`Scripting.Dictionary` は `CreateObject` のlate bindingで使う。
4. 判断: MsgBoxは使わず、`処理ログ` シートに処理結果、警告、エラーを一覧出力する。
5. 直接操作したファイル:
   - `experiments/001-claims-monthly-simple/prompt-records/codex-plan-medium.md`
   - `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/first-output.md`
   - `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/correction-request.md`
   - `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/final-output.md`
   - `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/final-vba.bas`
   - `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/notes.md`
6. 作成結果: 初回成果物として `first-output.md` を作成し、修正依頼未実施のため同内容を `final-output.md` に保存した。`final-vba.bas` は `final-output.md` 内のVBAコードブロックから無改変抽出した。
7. 確認質問なし/標準回答採用: 不明点は標準回答で足りたため、ユーザーへの追加確認質問は行わなかった。
8. 汚染防止: 他条件の `prompt-records/`、他条件の生成済みVBA本文、レビュー、評価、`verification/`、`reports/` は参照していない。書き込み先の存在確認時に `model-artifacts` 直下の条件ディレクトリ名のみ見えたが、他条件の内容は読まず、生成判断にも採用していない。

## 保存先

- first output: `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/first-output.md`
- 修正依頼記録: `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/correction-request.md`
- final output: `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/final-output.md`
- final VBA: `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/final-vba.bas`
- notes: `experiments/001-claims-monthly-simple/model-artifacts/codex-plan-medium/notes.md`

## 修正依頼

- 修正依頼回数: 0回
- 理由: Orchestrator検収でVBAコードの修正不要と判断。
- 対応: first output と同一内容を final output として扱う。コード改善・補完なし。

## 実験汚染の有無

- 他条件の成果物本文、レビュー、評価結果は参照していない。
- `verification/`、`reports/` は参照していない。
- 書き込み先確認で他条件ディレクトリ名が見えたが、中身は読んでいないため生成物へ採用していない。
- `AGENTS.md` と `README.md` は作業ルール、公開境界、リポジトリ入口確認として読んだ。モデル投入用素材としては扱っていない。
- `origin/main` 起点の #17 branch では #15/#16成果物は存在しなかったものとして扱い、生成判断に採用していない。
