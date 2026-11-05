---
name: excel-2010-compat
description: Use before releasing files to the user's local Excel 2010 environment.
---

# Make workbooks practical in Excel 2010

**Trigger:** Use before releasing files to the user's local Excel 2010 environment.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Scan modern functions, tables, charts and formatting.
2. replace unsupported dynamic arrays/modern functions.
3. check performance and normal zoom.
4. verify workbooks with desktop Excel 2010 where possible.

## Acceptance criteria
Document features requiring newer Excel and give 2010-compatible alternatives.

## Compatibility and failure modes
OOXML/ZIP validity alone does not prove Excel will open without a repair prompt.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
