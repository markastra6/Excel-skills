---
name: workbook-performance
description: Use for Excel 2010 not responding, long recalculation or oversized workbooks.
---

# Diagnose slow or freezing Excel files

**Trigger:** Use for Excel 2010 not responding, long recalculation or oversized workbooks.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Profile used ranges, volatile formulas, full-column array formulas, huge conditional formatting spans and unnecessary table objects.
2. isolate heavy sheets.
3. reduce formula repetition and unused styles.
4. compare calculation times on actual Excel.

## Acceptance criteria
Record observed improvements without guessing and preserve numerical parity.

## Compatibility and failure modes
Never replace working formulas with cached values without permission.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
