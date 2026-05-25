# explorer log

## 目的

日次クレーム集計ファイルから月次クレーム集計表へ、支店コード + 業務コード単位でクレーム件数を転記する標準モジュール用VBAを作る。

## 参照ファイル

- `experiments/001-claims-monthly-simple/generation-materials/task-spec.md`
- `experiments/001-claims-monthly-simple/generation-materials/prompt-template.md`
- `experiments/001-claims-monthly-simple/generation-materials/prompt-constraints.md`
- `experiments/001-claims-monthly-simple/generation-materials/conditions.md`
- `experiments/001-claims-monthly-simple/generation-materials/clarification-answer-bank.md`
- `experiments/001-claims-monthly-simple/samples/README.md`
- `experiments/001-claims-monthly-simple/samples/source/`
- `experiments/001-claims-monthly-simple/samples/expected/`
- `experiments/001-claims-monthly-simple/samples/workbooks/`

## 確認した仕様

- 対象月は `2605`。
- 日次ファイルは月次ファイルと同じフォルダ配下の `daily/2605/`。
- 日次シート `日次集計`: A 支店コード, B 支店名, C 業務コード, D 業務名, E クレーム件数。
- ファイル名 `YYMMDD` から日付列を判断する。
- 月次シート `月次集計`: A-D 支店/業務, E:AI 1日から31日, AJ 月合計。
- 支店コード+業務コード一致行へ転記する。

## 判断

- 転記前に日別列と月合計をクリアする。
- 同日同キー複数行は日次ファイル内で合算する。
- 月次表にないキーはログ記録してスキップする。
- 日次ファイル欠落、シート欠落、数値不正はログに残して処理を継続する。
- 日別列は固定仕様のE:AIを使用し、対象日はファイル名から列番号へ変換する。

## 見せなかった範囲

- 他条件の `prompt-records/`
- 他条件の `model-artifacts/`
- `verification/`
- `reports/`
- #15から#18の成果物、レビュー、生成済みVBA本文

## 公開境界確認

- Explorerは許可ファイルのみ参照した。
- 実データ、実パス、機密情報、非公開情報は扱っていない。
- サンプルの支店、業務、件数は公開可能な合成データとして扱った。
