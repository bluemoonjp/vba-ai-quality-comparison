# vba-ai-quality-comparison Agent Rules

## Scope

These rules apply to this independent repository. Parent `_Workspace` rules still apply where they do not conflict.

## Public Data Boundary

- Use only fictional, synthetic, publishable data.
- Do not add real customer, employee, branch, business, complaint, file path, workbook metadata, screenshots, secrets, or private source material.
- Treat `.xlsm` files as potentially sensitive until metadata and hidden content have been checked.
- Follow `docs/public-boundary.md` before adding any workbook, prompt log, generated output, screenshot, or exported VBA artifact.

## Experiment Integrity

- Compare AI usage patterns, not hidden human improvements.
- Keep the same task specification, sample data, expected results, and evaluation rubric across all conditions.
- Each condition gets at most one clarification question and one correction request.
- Record prompts, clarification questions, answers, correction requests, model/mode/date, output, and verification notes.
- Do not let later conditions see earlier generated solutions unless the experiment document explicitly marks the run as contaminated or learning-enabled.

## VBA And Excel Safety

- Store reviewable VBA as exported `.bas` / `.cls` text files.
- Do not run macros or write into existing workbooks without an explicit task and a clear target file.
- Windows Excel verification is required for completion of VBA behavior checks.
- Static review alone is not enough for final validation.

## Git And GitHub

- This directory is its own Git repository. Do not commit these files to the parent `_Workspace` repository.
- Keep commits scoped to one meaningful issue or phase.
- GitHub Issues are the task tracker for this repository.
- Record scaffold and public-safety audit updates in `docs/scaffold-audit-issue-1.md` or the relevant later issue document.
