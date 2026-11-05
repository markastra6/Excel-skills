---
name: pivot-summary
description: Use for grouping by staff, unit, stage, date, correction category and project.
---

# Summarize with PivotTables or formulas

**Trigger:** Use for grouping by staff, unit, stage, date, correction category and project.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Verify fact-table grain before aggregating.
2. define grouping fields and numeric measures.
3. avoid summing percentages.
4. check pivots refresh and hidden caches.
5. use SUMIFS/COUNTIFS-based summary when refresh-free is required.

## Acceptance criteria
Reconcile total counts and pages to source, with refresh instructions for pivots.

## Compatibility and failure modes
Excel 2010 pivot caches may not refresh automatically; don't promise otherwise.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
