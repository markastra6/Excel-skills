---
name: inspect-workbook
description: Use for unknown Excel files, hidden data and layout discovery.
---

# Analyze an existing workbook before edits

**Trigger:** Use for unknown Excel files, hidden data and layout discovery.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Inventory sheets and visibility.
2. identify merged cells, tables, named ranges, validation lists, charts and macros.
3. map input cells versus formulas.
4. determine source of truth.
5. inspect a representative section visually.
6. list broken links, mixed types and formula errors.

## Acceptance criteria
Deliver a sheet map, formula dependency summary, input/output map and confidence notes. Do not modify the original.

## Compatibility and failure modes
Do not infer that cached formula values are fresh; Excel 2010 calculation mode may be manual.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
