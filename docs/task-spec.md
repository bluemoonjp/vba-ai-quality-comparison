# VBA Task Specification

## Scenario

Daily complaint summaries are collected in subfolders as workbooks named `クレーム集計YYMMDD.xlsx`. A monthly workbook named `月次クレーム集計YYMM.xlsm` contains the VBA macro and monthly summary table.

The macro reads daily files for the target month and transfers complaint counts by branch and business line into the monthly table.

## Monthly Layout

- Row key: branch + business line.
- Columns: days `1` to `31` plus monthly total.
- The row-to-day mapping is the same for all rows.
- Subtotal rules are intentionally irregular:
  - Large branches have their own branch subtotal.
  - Smaller branches are grouped into regional subtotals.
  - Subtotal rows are separate from data rows and must not be treated as input keys.

## Daily Layout

The exact synthetic workbook layout will be defined in sample data work. It must include:

- branch
- business line
- complaint count
- report date derived from the filename

## Required Behavior

- Search the configured subfolder tree for matching daily files.
- Interpret `YYMMDD` and `YYMM` consistently.
- Transfer counts into the matching branch + business row and day column.
- Recalculate monthly totals and defined subtotals.
- Define behavior for missing days, duplicate day files, unknown branch/business pairs, blank counts, and non-numeric counts.
- Be safe to rerun according to the final sample specification.

## Out Of Scope

- Real operational data.
- Production deployment.
- Replacing the workbook design with another reporting system.
