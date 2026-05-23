# Samples

Synthetic workbook samples and expected results will be created here.

The sample set must include:

- daily workbook layout for `クレーム集計YYMMDD.xlsx`
- monthly workbook layout for `月次クレーム集計YYMM.xlsm`
- branch/business master
- large-branch and regional subtotal classification
- expected daily transfers, monthly totals, and subtotals
- abnormal cases for verification

Any workbook committed here must pass the public-safety checklist in the repository README.

## Source Files

- `source/branches.csv`: fictional branch and subtotal grouping master.
- `source/business-lines.csv`: fictional business line master.
- `source/daily-records.csv`: source rows used to generate daily workbook samples.

## Expected Results

- `expected/monthly-expected.csv`: expected monthly transfer results for valid source rows.
- `expected/anomaly-expected.csv`: expected handling for abnormal cases.

## Checked Workbooks

Checked `.xlsx` samples are stored under `checked/`.

- Daily samples are under `checked/daily/<YYMM>/`.
- Monthly layout samples are under `checked/monthly/`.
- `.xlsx` files are generated artifacts from the CSV source files.
- `.xlsm` files are intentionally deferred until the VBA implementation phase.

## Regeneration

Run `tools/build-samples.mjs` with the bundled Codex workspace Node runtime to regenerate expected CSV files and checked `.xlsx` samples.

By default, the script regenerates expected CSV files and creates any missing checked workbook files. Use `--force` only when intentionally replacing existing checked workbook files.

Use `--render` when a visual render pass is needed in an environment where artifact-tool rendering is available with a clean exit status.
