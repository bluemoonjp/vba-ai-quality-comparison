# Run Log Template

Copy this file to `outputs/<condition-id>/run-log.md` for each condition.

## Condition Metadata

- Condition id:
- Tool/model/mode:
- Reasoning setting, if visible:
- Run date:
- Operator:
- Task specification version:
- Sample data version:
- Evaluation rubric version:

## Common Prompt

Record the exact prompt sent to the condition.

## Clarification Question

- Asked by condition: yes/no
- Question:

## Clarification Answer

Record the exact answer sent. If no question was asked, write `Not used`.

## First Output

Record or link to the unedited first output.

## Verification Before Correction

- Verification checklist used:
- Key defects found:
- Not-run items and reasons:

## Correction Request

Record the exact single correction request sent. If no correction was possible, explain why.

## Final Output

Record or link to the unedited final output.

## Storage And Conversion Notes

Record any export, import, formatting, or file-normalization step. Do not silently fix generated code.

## Contamination Check

- Was any prior condition output visible to this condition?
- Was any extra hint given beyond the protocol?
- Notes:

## Public-Safety Check

- Contains only synthetic/public data:
- Contains no real paths or private identifiers:
- Contains no secrets or private connector output:

## Open Questions

Record unresolved assumptions or missing evidence.
