---
name: reconcile-compare
description: Use for before/after workbooks, source vs tracker, monthly revisions and data migrations.
---

# Compare workbook versions or source records

**Trigger:** Use for before/after workbooks, source vs tracker, monthly revisions and data migrations.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Build workbook manifests.
2. compare sheet names, formula counts, charts, cell values and formatting indicators.
3. align rows by stable IDs, not row position.
4. classify expected vs unexpected differences.
5. check totals.

## Acceptance criteria
Produce changed/missing/extra summaries and protect source files.

## Compatibility and failure modes
Do not assume formula cached results are current or identical across calculation engines.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
