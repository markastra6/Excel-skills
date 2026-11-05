---
name: migration-modernization
description: Use for .xls to .xlsx, Excel 2010 to newer Excel, or old formatting/templates.
---

# Upgrade legacy Excel without data loss

**Trigger:** Use for .xls to .xlsx, Excel 2010 to newer Excel, or old formatting/templates.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Inventory legacy features first.
2. define target feature compatibility.
3. preserve names, charts, validation and formulas.
4. compare source and converted totals.
5. treat macro-enabled files separately.
6. retain an original copy.

## Acceptance criteria
Log unsupported or changed objects and perform target-version open tests.

## Compatibility and failure modes
Conversion may change cached values or remove unsupported legacy features; never silently overwrite.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
