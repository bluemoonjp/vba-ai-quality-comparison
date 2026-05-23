# vba-ai-quality-comparison

VBA program generation experiments for comparing output quality across different AI usage patterns.

This repository is intentionally public. It uses only fictional branch, business, folder, and complaint-count data. Do not add real business data, personal data, internal paths, secrets, or screenshots derived from real workbooks.

## Purpose

The experiment compares the quality of VBA deliverables produced under these conditions:

- ChatGPT Instant
- ChatGPT Thinking extended
- Codex, simple plan-to-implementation flow, medium reasoning
- Codex, simple plan-to-implementation flow, very high reasoning
- Codex, multi-agent review flow
  - Orchestrator: high
  - Explorer: medium
  - Reviewer: high

The task is to create a macro embedded in a monthly workbook named `月次クレーム集計YYMM.xlsm`. The macro reads daily workbooks named `クレーム集計YYMMDD.xlsx` from subfolders and transfers complaint counts by branch and business line into the monthly summary.

## Current State

Initial repository scaffold. The experiment is managed through GitHub Issues so that each phase can be completed and reviewed independently.

## Experiment Protocol

- All conditions use the same task specification, sample data, expected results, and evaluation rubric.
- Each condition may ask one clarification question and receive one follow-up correction request after the first output.
- Human edits are not allowed before evaluation. Format conversion and exported VBA storage are allowed only when logged.
- Model, mode, date, prompt, clarification, answer, correction request, final output, and verification result must be recorded.

## Repository Layout

- `docs/`: experiment design, task specification, evaluation rubric, and final summary.
- `prompts/`: common and condition-specific prompts.
- `outputs/`: condition-specific generation logs and final outputs.
- `samples/`: synthetic workbook specifications, fixtures, and expected results.
- `src/`: exported VBA modules, stored as text-first review artifacts.
- `tests/`: manual Excel verification scenarios and result logs.
- `.github/issues/`: issue body drafts used for initial GitHub issue creation.

## Public Boundary

See `docs/public-boundary.md` before adding prompts, logs, workbooks, exported VBA, screenshots, or generated outputs. Issue #1 scaffold checks are recorded in `docs/scaffold-audit-issue-1.md`.

## Git Boundary

This directory is an independent Git repository under the parent `_Workspace` folder. Do not mix commits or issue work between this repository and the parent workspace repository.

## Public Safety Checklist

Before pushing or publishing generated artifacts, confirm:

- No real branch, customer, employee, company, or complaint details are present.
- No real folder paths, usernames, machine names, or OneDrive paths are embedded.
- No workbook hidden sheets, document properties, comments, or metadata contain private information.
- No secrets, tokens, API responses, screenshots, or private source material are included.
- VBA source is exported to text files for review; `.xlsm` files are treated as verification samples.

## License

MIT License. See `LICENSE`.
