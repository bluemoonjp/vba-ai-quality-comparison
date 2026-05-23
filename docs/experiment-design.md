# Experiment Design

## Goal

Compare VBA deliverable quality across different AI usage patterns using one shared business-style automation task.

## Conditions

| condition id | condition | allowed interaction |
| --- | --- | --- |
| `chatgpt-instant` | ChatGPT Instant | one clarification question, one correction request |
| `chatgpt-thinking-extended` | ChatGPT Thinking extended | one clarification question, one correction request |
| `codex-medium` | Codex simple plan-to-implementation, medium reasoning | one clarification question, one correction request |
| `codex-xhigh` | Codex simple plan-to-implementation, very high reasoning | one clarification question, one correction request |
| `codex-multi-agent-review` | Codex orchestrator high + explorer medium + reviewer high | one clarification question, one correction request |

These condition ids are the stable folder, log, and report identifiers for the experiment.

## Execution Protocol

Each condition uses the same execution sequence:

1. Provide the common task prompt and the same task specification.
2. Allow at most one clarification question from the AI condition.
3. Answer the clarification question with the approved standard answer set and the same level of detail for every condition.
4. Save the first output without human code edits.
5. Apply the same verification checklist to identify correction points.
6. Send at most one correction request.
7. Save the final output without human code edits.
8. Export or normalize files only for storage, and log every conversion step.
9. Evaluate the final output with the shared rubric and verification evidence.

If a condition does not ask a clarification question, the unused clarification turn is not replaced by extra instructions. If a condition cannot produce runnable or reviewable output, record the failure and continue to evaluation.

## Fairness Rules

- Use the same task specification, samples, expected results, and rubric for all conditions.
- Give clarification answers at the same level of detail.
- Base correction requests on the same verification checklist.
- Do not manually improve generated VBA before scoring.
- Log any format conversion, export/import step, or failed execution.
- Do not expose earlier condition outputs, reviews, fixes, or scores to later conditions.
- Do not improve the prompt between conditions unless the entire experiment is restarted or the run is explicitly marked as contaminated.
- Use `outputs/run-log-template.md` for condition logs and `tests/evaluation-sheet-template.md` for evaluation records.

## Human Intervention Boundary

Allowed human actions:

- paste the fixed common prompt into each condition
- answer the single clarification question using approved information
- send the single correction request
- save, export, or format generated output for repository storage
- run verification and record evidence

Disallowed human actions before scoring:

- rewrite generated VBA
- add missing logic
- silently correct syntax or object model mistakes
- give one condition extra hints that other conditions did not receive
- reuse a better solution from another condition

## Public Boundary

All workbooks, branch names, business names, folder names, and counts are synthetic. Public release requires metadata and hidden-content checks for any workbook.
