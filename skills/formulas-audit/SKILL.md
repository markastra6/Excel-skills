---
name: formulas-audit
description: Use for #REF!, #VALUE!, wrong totals, hardcoded rates, circularity and unexpected cached results.
---

# Trace and correct formula failures

**Trigger:** Use for #REF!, #VALUE!, wrong totals, hardcoded rates, circularity and unexpected cached results.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Identify source and downstream dependent cells.
2. compare formulas across rows and columns.
3. inspect relative references, ranges, text-number mismatches and named ranges.
4. test zero, missing and overplan data.
5. check calculation settings.

## Acceptance criteria
Provide root cause, fixed formula and before/after test evidence, without silently rewriting unrelated logic.

## Compatibility and failure modes
Do not equate empty formula caches with failed formulas; recalculate in desktop Excel.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
