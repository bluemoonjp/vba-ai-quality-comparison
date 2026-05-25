# Excel実行ログ

## 環境

- OS: Windows
- Excel version: 16.0
- 実行日: 2026-05-26
- 実行者: Orchestrator / Worker

## 手順

1. サンプル月次workbookをscratchへコピーする。
2. 日次workbook `daily/2605` を同じscratchへコピーする。
3. Excel VBE import用に、条件別 `final-vba.bas` のCP932変換コピーをscratch内に作成する。
4. scratchを一時ドライブへ割り当て、Excelからローカルパスとして月次workbookを開く。
5. 月次workbookをscratch内で一時 `.xlsm` として保存し、条件別VBAをimportする。
6. 入口プロシージャを実行する。
7. `samples/expected/monthly-expected.csv` と結果を比較する。
8. 同じworkbookセッションで再実行し、二重加算がないことを確認する。

## 実行時の補足

- 追跡対象の `.bas` は評価前に修正していない。
- 初回試行では、UTF-8の `.bas` をそのままVBEへimportしたため日本語文字列が文字化けし、コンパイルエラーのダイアログで停止した。以後はscratch内のCP932変換コピーを使った。
- cloud-synced folder配下のscratchを直接Excelで開くと `ThisWorkbook.Path` がURLになり、VBAの `Dir()` が日次フォルダを検出できなかった。以後は同じscratchを一時ドライブへ割り当てて実行した。
- scratch workbookとimport用コピーは成果物として保存しない。

## 結果

| condition_id | entry_procedure | import | run | expected_match | rerun_safe | no_msgbox | major_defect | notes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| codex-chatgpt-medium | `ImportDailyClaimsToMonthly` | pass | pass | pass | pass | pass | no | `branch_code + business_code` キーで期待値と一致。その他日別列も0。 |
| codex-chatgpt-xhigh | `ImportDailyClaimsToMonthly` | pass | pass | pass | pass | pass | no | `branch_code + business_code` キーで期待値と一致。その他日別列も0。 |
| codex-plan-medium | `RunClaimsMonthlyTransfer` | pass | pass | pass | pass | pass | no | `branch_code + business_code` キーで期待値と一致。その他日別列も0。 |
| codex-plan-xhigh | `ImportDailyClaimsToMonthly` | pass | pass | pass | pass | pass | no | `branch_code + business_code` キーで期待値と一致。その他日別列も0。 |
| codex-multi-review | `TransferDailyClaimsToMonthly` | pass | pass | pass | pass | pass | no | `branch_code + business_code` キーで期待値と一致。その他日別列も0。 |

## 静的確認

- `MsgBox` / `InputBox`: 検出なし。
- 固定実パス: 検出なし。
- 外部通信: 検出なし。
- 追加参照: 必須指定なし。`Scripting.Dictionary` はlate bindingで使用。
