---
name: edit-preserve
description: Use for edits where formulas, metadata, staff data, formatting and charts must survive.
---

# Safely modify an existing workbook

**Trigger:** Use for edits where formulas, metadata, staff data, formatting and charts must survive.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Make a copy.
2. record starting sheet names and counts of formulas, charts, defined names, validations, conditional formats and merges.
3. scope smallest changes.
4. do not silently normalize original data.
5. compare before/after snapshots.

## Acceptance criteria
Report exact cells or sheets changed; prove formulas and objects were preserved unless a change was requested.

## Compatibility and failure modes
Never round-trip macro-enabled .xlsm through a tool that cannot preserve VBA.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
