---
name: testing-release
description: Use before handing over a workbook or pushing a template to GitHub.
---

# Test and ship reliable Excel workbooks

**Trigger:** Use before handing over a workbook or pushing a template to GitHub.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Run ZIP/XML audit, manifest comparison and formula-error scan.
2. exercise blank/zero/negative/overplan/new staff/rate changes.
3. verify charts, validations, freeze panes and open/reopen.
4. test in Excel 2010 where possible.

## Acceptance criteria
Provide a checklist with passed, failed and untested claims, and a verified output path.

## Compatibility and failure modes
Never claim Excel desktop compatibility based only on a Python reader.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
