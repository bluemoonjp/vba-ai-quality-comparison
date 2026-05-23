# Sample Data Specification

Issue #3 defines the synthetic sample set used by every AI condition.

## Scope

The sample data is fictional and public-safe. It covers three target months:

- `2026-02`: 28-day month
- `2026-04`: 30-day month
- `2026-05`: 31-day month

The monthly workbook layout always includes day columns `1` through `31` plus a monthly total. Days outside the target month remain blank in expected results.

## Source Of Truth

- `samples/source/branches.csv`: branch master and subtotal grouping.
- `samples/source/business-lines.csv`: business line master.
- `samples/source/daily-records.csv`: synthetic rows used to build daily workbooks.
- `samples/expected/monthly-expected.csv`: expected monthly transfer results generated from valid source rows.
- `samples/expected/anomaly-expected.csv`: expected handling for abnormal sample cases.

Checked `.xlsx` workbooks in `samples/checked/` are regenerated artifacts, not the primary source of truth.

## Branch Master

The sample uses six fictional branches across three fictional regions.

- Large branches use their own subtotal group.
- Small branches are subtotaled by region.
- Branch and region names intentionally avoid real-world naming.

## Business Lines

The sample uses eight fictional business lines. Each branch supports only some business lines, so the branch + business matrix is intentionally uneven.

## Daily Workbook Layout

Each daily workbook is named `クレーム集計YYMMDD.xlsx` and contains one sheet named `日次集計`.

Columns:

1. `処理区分`
2. `支店コード`
3. `支店名`
4. `業務コード`
5. `業務名`
6. `クレーム件数`
7. `備考`

Rows marked `ok` are valid source rows. Other `処理区分` values represent abnormal cases and are documented in `samples/expected/anomaly-expected.csv`.

## Monthly Workbook Layout

Each monthly workbook is named `月次クレーム集計YYMM.xlsx`. It is an `.xlsx` layout sample for the future macro-enabled `月次クレーム集計YYMM.xlsm`.

Sheets:

- `月次集計`: blank target layout with row keys, day columns, formulas for monthly totals and subtotals.
- `期待結果`: expected populated result for valid source rows.
- `マスタ`: branch/business and subtotal reference tables.
- `異常系`: expected abnormal-case handling for that month.

Rows in monthly sheets include:

- `DETAIL`: branch + business rows used for transfers.
- `SUBTOTAL`: subtotal rows that must not be treated as input keys.

## Abnormal Cases

The sample includes:

- missing daily file days
- duplicate day file in another subfolder
- unknown branch
- unknown business line
- blank complaint count
- non-numeric complaint count
- out-of-month day columns in 28-day and 30-day months

Expected behavior for abnormal source rows is to skip transfer and record/log the issue. Missing days and out-of-month columns remain blank in expected results.

## Deferred Work

This issue does not create `.xlsm` files, VBA modules, or macro behavior. Those are handled by later issues after the AI-facing task specification is finalized.
