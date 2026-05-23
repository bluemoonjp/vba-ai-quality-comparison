# Evaluation Rubric

Each AI condition is evaluated with the same checklist and evidence.

## Scoring Scale

Each category is scored from `0` to `4`.

| score | meaning |
| --- | --- |
| 0 | Missing, unusable, or unsafe for the category. |
| 1 | Major defects; works only in narrow or accidental cases. |
| 2 | Partially adequate; important gaps remain. |
| 3 | Good; minor gaps or risks remain. |
| 4 | Strong; meets the expected behavior with clear evidence. |

Every score must include a short evidence note. Do not use the total score alone as the conclusion; also record severe defects, failure patterns, and correction-request impact.

## Categories

| category | focus |
| --- | --- |
| Correctness | Transfers daily counts to the correct branch, business, and day cells. |
| Subtotals | Handles large-branch and regional subtotal rules correctly. |
| Rerun safety | Produces predictable results when run more than once. |
| Error handling | Handles missing files, unknown keys, blank counts, non-numeric counts, and filename mismatches. |
| Maintainability | Uses clear procedures, names, constants, and separable responsibilities. |
| Excel/VBA fit | Avoids fragile active-object usage and restores Excel application state. |
| Performance | Reads workbook data without unnecessary cell-by-cell overhead where practical. |
| Explanation quality | Explains assumptions, usage, limitations, and test approach. |

## Required Evaluation Notes

For each condition, record:

- category scores and evidence
- not-run items and reasons
- severe defects that would block practical use
- whether the single correction request improved the output
- any suspected contamination or unequal information exposure
