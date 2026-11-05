---
name: formulas-engineering
description: Use for automatic totals, allocation rules, date math, quality metrics and reusable target controls.
---

# Build robust formulas and calculation logic

**Trigger:** Use for automatic totals, allocation rules, date math, quality metrics and reusable target controls.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Write inputs and units first.
2. centralize the single source of targets.
3. use relative/absolute references deliberately.
4. protect division by zero.
5. blank when data is unknown rather than zero.
6. guard inconsistent inputs.
7. test examples and boundary cases.

## Acceptance criteria
Record each calculated field, its definition, references, and expected test results.

## Compatibility and failure modes
Prefer SUMIFS, COUNTIFS, INDEX/MATCH, IFERROR when supported; avoid XLOOKUP, FILTER, LET and dynamic arrays for Excel 2010.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
