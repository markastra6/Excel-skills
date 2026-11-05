---
name: tables-validation
description: Use when staff, units, categories or dates must be extendable.
---

# Set up editable tables and dropdowns

**Trigger:** Use when staff, units, categories or dates must be extendable.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Use unambiguous headers and clear input styling.
2. validate allowable status/work types.
3. avoid formulas ending at a fixed short range.
4. add checks for missing required fields and negative outputs.
5. test inserted rows.

## Acceptance criteria
Confirm adding a staff member, chapter and date does not break calculation coverage.

## Compatibility and failure modes
Excel Tables exist in 2010, but structured-reference and chart behavior must be tested.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
