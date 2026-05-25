# 評価表

| condition_id | import | run | transfer | monthly_total | rerun_safe | no_msgbox | maintainability | major_defect | notes |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| codex-chatgpt-medium | pass | pass | pass | pass | pass | pass | 未評価 | no | `ImportDailyClaimsToMonthly`。期待値と一致。再実行でも二重加算なし。 |
| codex-chatgpt-xhigh | pass | pass | pass | pass | pass | pass | 未評価 | no | `ImportDailyClaimsToMonthly`。期待値と一致。再実行でも二重加算なし。 |
| codex-plan-medium | pass | pass | pass | pass | pass | pass | 未評価 | no | `RunClaimsMonthlyTransfer`。期待値と一致。再実行でも二重加算なし。 |
| codex-plan-xhigh | pass | pass | pass | pass | pass | pass | 未評価 | no | `ImportDailyClaimsToMonthly`。期待値と一致。再実行でも二重加算なし。 |
| codex-multi-review | pass | pass | pass | pass | pass | pass | 未評価 | no | `TransferDailyClaimsToMonthly`。期待値と一致。再実行でも二重加算なし。 |

## 判定メモ

- `pass`: 条件を満たした。
- `fail`: 条件を満たさない。
- `blocked`: 検証不能。
- `未実施`: まだ検証していない。
- `未評価`: 今回の実Excel実行では採点しない。保守性比較は後続レポートで扱う。

## 評価メモ

- 実行日: 2026-05-26。
- 実行環境: Windows版Excel 16.0。
- `final-vba.bas` は評価前に修正していない。
- Excel VBE importでは、scratch内にCP932変換したimport用コピーを作成して使用した。追跡対象の `.bas` は変更していない。
- cloud-synced folder配下のworkbookを直接開くと `ThisWorkbook.Path` がURLになるため、scratchを一時ドライブへ割り当ててExcelからローカルパスとして開いた。
- 期待値照合は `branch_code + business_code` をキーにし、`day_01`, `day_02`, `day_15`, `month_total` と、その他日別列が0のままであることを確認した。
