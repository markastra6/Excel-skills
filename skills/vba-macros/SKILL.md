---
name: vba-macros
description: Use for buttons, imports, user forms, workbook events or repeatable local actions.
---

# Engineer safe Excel VBA automation

**Trigger:** Use for buttons, imports, user forms, workbook events or repeatable local actions.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Inventory macro dependencies.
2. preserve signed originals.
3. write Option Explicit, bounded ranges and explicit error handling.
4. disable and restore events/screen updating/calculation in Finally-equivalent cleanup.
5. log steps and protect open files.

## Acceptance criteria
Deliver .xlsm code and clear macro enable instructions; test interruption and repeated runs.

## Compatibility and failure modes
Do not put secrets in VBA or claim Excel 2010 macro compatibility without desktop tests.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
