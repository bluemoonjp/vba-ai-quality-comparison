# Experiment Design

## Goal

Compare VBA deliverable quality across different AI usage patterns using one shared business-style automation task.

## Conditions

| id | condition | allowed interaction |
| --- | --- | --- |
| `chatgpt-instant` | ChatGPT Instant | one clarification question, one correction request |
| `chatgpt-thinking-extended` | ChatGPT Thinking extended | one clarification question, one correction request |
| `codex-medium` | Codex simple plan-to-implementation, medium reasoning | one clarification question, one correction request |
| `codex-xhigh` | Codex simple plan-to-implementation, very high reasoning | one clarification question, one correction request |
| `codex-multi-agent-review` | Codex orchestrator high + explorer medium + reviewer high | one clarification question, one correction request |

## Fairness Rules

- Use the same task specification, samples, expected results, and rubric for all conditions.
- Give clarification answers at the same level of detail.
- Base correction requests on the same verification checklist.
- Do not manually improve generated VBA before scoring.
- Log any format conversion, export/import step, or failed execution.

## Public Boundary

All workbooks, branch names, business names, folder names, and counts are synthetic. Public release requires metadata and hidden-content checks for any workbook.
