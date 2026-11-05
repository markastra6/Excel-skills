---
name: production-qc
description: Use for manuscript, proof, JA, composition, QC, correction rounds and deliveries.
---

# Track production and QC per chapter or unit

**Trigger:** Use for manuscript, proof, JA, composition, QC, correction rounds and deliveries.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Store manuscript pages and proof pages separately.
2. group SB/WS/SA.
3. allocate Comp and QC owners.
4. centralize rates by work type.
5. calculate pages done, chapter-wise left, over-plan and status.
6. distinguish correction count and confirmed errors.

## Acceptance criteria
Preserve actual source figures; reconcile chapter, staff and day totals without inventing missing allocations.

## Compatibility and failure modes
Multiple assignees need explicit allocation rules; concatenated names are not sufficient to infer each person's pages.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
