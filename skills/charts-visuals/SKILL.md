---
name: charts-visuals
description: Use for planned vs actual, chapter backlog, daily trends, and error distribution.
---

# Build clear Excel charts and sparklines

**Trigger:** Use for planned vs actual, chapter backlog, daily trends, and error distribution.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Select measurable questions.
2. validate category labels and chart ranges.
3. use columns for categorical comparisons, lines for trends, bars for ranked work.
4. exclude or label missing data.
5. place charts without hiding inputs.

## Acceptance criteria
Charts must update for the intended input range, avoid overlap and identify units and periods.

## Compatibility and failure modes
Verify rendered Excel 2010 charts; never infer results from chart artwork alone.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
