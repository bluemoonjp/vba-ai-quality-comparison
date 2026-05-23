# Public Boundary

This repository is public by design. Every artifact must be safe to publish before it is committed.

## Allowed Artifacts

- Synthetic workbook specifications, synthetic master data, and expected results.
- Exported VBA source files such as `.bas` and `.cls`.
- Prompts, generation logs, and review notes that have been sanitized for public release.
- Sample workbooks only after the workbook checklist below is complete.
- Documentation, rubrics, test cases, and comparison summaries that use fictional data.

## Disallowed Artifacts

- Real operational data, complaint contents, branch names, customer names, employee names, or company-specific identifiers.
- Real local paths, user names, machine names, OneDrive paths, network paths, or internal file names.
- Secrets, tokens, API responses, private connector output, or private source material.
- Private screenshots or screenshots derived from real workbooks.
- Unchecked Office binaries, including `.xls`, `.xlsx`, `.xlsm`, and `.xlsb`.
- Workbook content that may hide private data, such as hidden sheets, comments, names, links, or document properties.

## Workbook Publication Checklist

Before any workbook is moved into `samples/checked/` and committed, verify:

- The workbook uses fictional data only.
- Document properties and author metadata are cleared or intentionally public.
- Hidden and very hidden sheets are reviewed.
- Comments, notes, threaded comments, and cell metadata are reviewed.
- Defined names, formulas, external links, pivot caches, connections, and queries do not contain real paths or private data.
- VBA modules, forms, references, and constants do not contain real paths, private names, or secrets.
- The workbook opens in Windows Excel and the intended macro/security behavior is documented.
- The file name itself is safe to publish.

## Git Boundary

This directory is an independent Git repository under the parent workspace. Commits, remotes, tags, and GitHub Issues belong to this repository unless explicitly stated otherwise.

Do not commit this repository's files into the parent workspace repository. Do not copy private parent workspace context into this repository.

## Default Storage Rule

Unchecked generated workbooks belong in ignored local scratch locations such as `samples/generated/` or `outputs/<condition>/scratch/`. Checked public workbooks may be placed in `samples/checked/` after the checklist is complete.
