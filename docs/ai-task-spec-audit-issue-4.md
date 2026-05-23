# AI Task Specification Audit For Issue #4

Audit date: 2026-05-23

## Scope

Issue #4 fixes the task specification and common prompt that will be given to every AI condition. It does not execute any AI condition, generate VBA, edit sample workbooks, or create `.xlsm` files.

## Decisions Fixed

- Prompt language: Japanese.
- Experiment interaction limit: 仕様確認1回 and 修正依頼1回.
- Expected code: VBA for a standard module, preferably with readable English identifiers.
- Explanation language: Japanese.
- Search root: `ThisWorkbook.Path\daily\YYMM\`.
- Target month source: monthly workbook file name `月次クレーム集計YYMM.xlsm`.
- Daily date source: daily workbook file name `クレーム集計YYMMDD.xlsx`.
- Transfer key: branch code + business code.
- Row rules: `DETAIL` rows are transfer targets; `SUBTOTAL` rows are never input keys.
- Rerun behavior: clear detail day cells first, then rebuild results.
- Reference settings: AI may choose, but must explain any additional reference and an alternative.

## Artifacts Updated

- `docs/task-spec.md`: detailed AI-facing task specification.
- `prompts/common-task-prompt.md`: copy-ready common prompt for all AI conditions.
- `prompts/clarification-answer-bank.md`: standard answers for the single allowed clarification question.

## Handoff

- Issue #5 should use `prompts/common-task-prompt.md` as the initial prompt for every AI condition.
- If a condition asks a clarification question, answer from `prompts/clarification-answer-bank.md`.
- Do not give extra implementation hints outside the prompt and answer bank.
- Keep generated outputs unedited before scoring.

## Deferred Work

- Actual AI condition execution is deferred to Issue #5.
- Windows Excel verification is deferred to Issue #6.
- Macro-enabled `.xlsm` artifacts remain deferred until a VBA implementation is selected or tested.
