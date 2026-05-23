# Protocol Audit For Issue #2

Audit date: 2026-05-23

## Scope

Issue #2 fixes the experiment protocol, logging format, and evaluation scoring method. It does not run any AI condition, create Excel samples, or generate VBA.

## Decisions Fixed

- The experiment uses five stable condition ids:
  - `chatgpt-instant`
  - `chatgpt-thinking-extended`
  - `codex-medium`
  - `codex-xhigh`
  - `codex-multi-agent-review`
- Each condition may use at most one clarification question and one correction request.
- Scoring uses `0-4` points per category plus mandatory evidence notes.
- Total score is supporting information, not the sole conclusion.
- Severe defects, not-run checks, correction impact, and fairness/contamination notes must be recorded.

## Templates Added

- `outputs/run-log-template.md`
- `tests/evaluation-sheet-template.md`

These templates are the default starting points for Issue #5 generation logs and Issue #6 verification/evaluation records.

## Deferred Work

- Common task prompt finalization remains dependent on Issue #4.
- Synthetic workbook/sample creation remains dependent on Issue #3.
- Actual condition execution remains dependent on Issue #5.
- Windows Excel verification remains dependent on Issue #6.

## Closeout Criteria

Issue #2 is complete when:

- The protocol can be followed without choosing additional rules.
- The scoring scale and category evidence requirements are documented.
- The run-log and evaluation templates are available.
- No AI execution, Excel sample, or VBA implementation is introduced by this issue.
