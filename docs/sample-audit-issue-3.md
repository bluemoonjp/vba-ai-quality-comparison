# Sample Audit For Issue #3

Audit date: 2026-05-23

## Scope

Issue #3 creates the synthetic sample specification, CSV source data, expected results, and checked `.xlsx` workbooks. It does not create `.xlsm` files, VBA modules, prompts for AI execution, or generated AI outputs.

## Decisions Fixed

- Target months:
  - `2026-02`: 28-day month
  - `2026-04`: 30-day month
  - `2026-05`: 31-day month
- Source of truth:
  - `samples/source/*.csv`
  - `samples/expected/*.csv`
  - `docs/sample-data-spec.md`
- Checked `.xlsx` files are regenerated artifacts under `samples/checked/`.
- `.xlsm` files are deferred until the VBA implementation phase.

## Generated Artifacts

- Daily workbooks: 11 files under `samples/checked/daily/`.
- Monthly workbooks: 3 files under `samples/checked/monthly/`.
- Expected CSV files:
  - `samples/expected/monthly-expected.csv`
  - `samples/expected/anomaly-expected.csv`

## Verification Summary

- `tools/build-samples.mjs` generated the sample workbooks and expected CSV files.
- Monthly workbooks were inspected after generation for key `期待結果` ranges.
- Formula/error scan found no `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, or `#N/A` matches.
- Workbook render checks were attempted during verification. In this local environment the render command produced inspection output but returned a non-zero shell status, so stable closeout evidence relies on range inspection and formula/error scans.
- Repository scan found no Office binaries outside `samples/checked/`.
- Secret/path scan hits were limited to public-safety policy text.

## Public-Safety Notes

- All branch, region, business, and folder names are fictional.
- Workbooks are `.xlsx`, not `.xlsm`.
- The committed workbooks are sample layouts and expected-result artifacts, not operational data.

## Handoff

- Issue #4 should use `docs/sample-data-spec.md`, `samples/source/*.csv`, and `samples/expected/*.csv` to write the AI-facing task specification.
- Issue #5 should not alter these samples while collecting AI outputs.
- Issue #6 should use the checked `.xlsx` workbooks and expected CSV files for verification.
