---
name: conditional-formatting
description: Use for targets, performance, overdue work, status and quality warnings.
---

# Use colour to flag exceptions clearly

**Trigger:** Use for targets, performance, overdue work, status and quality warnings.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Choose meaning-first thresholds.
2. apply simple formats to limited ranges.
3. flag Done with balance >0, actual above plan and invalid PE count.
4. keep labels meaningful when printed monochrome.

## Acceptance criteria
Test each threshold boundary and confirm formatting follows inserted or copied rows.

## Compatibility and failure modes
Prefer native simple rules/data bars supported in Excel 2010.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
