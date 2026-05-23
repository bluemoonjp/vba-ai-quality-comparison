# VBA Task Specification

This document is the AI-facing task specification source. `prompts/common-task-prompt.md` wraps this content as the prompt used for every AI condition.

## Goal

Create an Excel VBA macro for a monthly complaint summary workbook.

The macro is stored in the monthly workbook and transfers daily complaint counts from daily workbooks into the monthly table by branch and business line.

## Target Workbook

- Macro host workbook name pattern: `月次クレーム集計YYMM.xlsm`
- Sample layout workbook name pattern: `月次クレーム集計YYMM.xlsx`
- Target months in the sample set:
  - `2602` = February 2026
  - `2604` = April 2026
  - `2605` = May 2026
- The macro must determine the target month from the monthly workbook file name.
- If the workbook name does not match the expected pattern, the macro must stop with a clear message.

## Daily File Search

- Daily workbook name pattern: `クレーム集計YYMMDD.xlsx`
- Search root: `ThisWorkbook.Path\daily\YYMM\`
- The macro must recursively search under that search root.
- Only files whose `YYMM` matches the target month are in scope.
- The report day is determined from the file name, not from worksheet cell values.
- Missing daily files are allowed. The corresponding day columns remain blank.
- If more than one daily file exists for the same `YYMMDD`, process only one file and log the duplicate. Duplicate rows from the skipped file must not be transferred.

## Daily Workbook Layout

Each daily workbook contains a sheet named `日次集計`.

Columns:

| column | header | role |
| --- | --- | --- |
| A | `処理区分` | sample scenario marker; `ok` rows are normal transfer rows |
| B | `支店コード` | branch key |
| C | `支店名` | display/check value |
| D | `業務コード` | business key |
| E | `業務名` | display/check value |
| F | `クレーム件数` | complaint count |
| G | `備考` | sample note |

The macro must transfer rows where `処理区分 = "ok"` and the branch/business/count values are valid. Other values are abnormal sample cases and must be logged/skipped.

## Monthly Workbook Layout

The monthly workbook contains these sheets:

- `月次集計`: target transfer sheet.
- `マスタ`: branch/business and subtotal reference tables.
- `異常系`: optional sheet for expected abnormal cases and/or log comparison.
- `期待結果`: expected result sheet in sample `.xlsx` workbooks; not required in a production `.xlsm`.

`月次集計` columns:

| column range | meaning |
| --- | --- |
| A | row type |
| B | subtotal group |
| C | region |
| D | branch code |
| E | branch name |
| F | business code |
| G | business name |
| H:AL | day columns `1` to `31` |
| AM | monthly total |
| AN | note |

Rows:

- `DETAIL`: valid transfer target rows. Use `支店コード + 業務コード` as the transfer key.
- `SUBTOTAL`: subtotal rows. These rows must never be treated as input keys.

Branch names and business names are display/check values. The primary matching key is branch code + business code.

## Transfer Rules

- Before transfer, clear only the day-count cells for `DETAIL` rows in columns `H:AL`.
- Do not clear row labels, master data, formulas, subtotal row labels, or unrelated sheets.
- For each valid daily row, add the complaint count to the matching day column for the matching `DETAIL` row.
- If multiple valid rows in the same daily file have the same branch/business key, add them together.
- Recalculate or restore monthly totals and subtotal rows after transfer.
- Days outside the target month remain blank and are excluded from monthly totals.
- The macro must be safe to rerun: running it twice with the same source files should produce the same result, not doubled counts.

## Subtotal Rules

Subtotal rules are intentionally irregular.

- Large branches have their own subtotal group.
- Small branches are subtotaled by region.
- `SUBTOTAL` rows are calculated from `DETAIL` rows with the same subtotal group.
- Subtotal rows cover the same day columns and monthly total column as detail rows.
- The macro may calculate subtotal values directly or preserve/fill formulas, but the final values must match the expected result.

## Error And Log Rules

The macro must skip transfer and record/log these cases:

- unknown branch code
- unknown business code
- missing branch/business target row
- blank complaint count
- non-numeric complaint count
- duplicate daily file for the same date
- daily file name that does not match `クレーム集計YYMMDD.xlsx`
- daily file whose `YYMM` does not match the target month
- missing `日次集計` sheet

The log can be an in-workbook sheet, message summary, or structured output table. The implementation must explain where the log is written and what columns it contains.

## Reference Setting Policy

The AI may decide whether to use additional VBA reference settings.

If additional references are used, the answer must explain:

- the reference name
- why it is useful
- what setup is required in Excel/VBE
- a late-binding or standard-VBA alternative

Using no additional references is acceptable.

## Required Output From AI

The AI response must include:

- complete VBA code suitable for a standard module
- the macro entry point name
- where to place the code
- how to run it
- whether any reference settings are required
- assumptions and limitations
- error/log handling explanation
- suggested test cases using the sample months `2602`, `2604`, and `2605`

Prefer readable English identifiers in VBA code. Japanese comments or user-facing messages are acceptable.

## Out Of Scope

- Running the macro.
- Creating or editing actual `.xlsm` files.
- Changing the sample workbook layout.
- Using real operational data.
- Production deployment.
