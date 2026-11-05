---
name: data-cleaning
description: Use for inconsistent text, dates, blanks, duplicates and import issues.
---

# Clean and standardize Excel data

**Trigger:** Use for inconsistent text, dates, blanks, duplicates and import issues.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Preserve raw inputs.
2. profile nulls, type mismatches and whitespace.
3. normalize line breaks, Unicode and date formats.
4. track transformations.
5. deduplicate by approved composite keys only.
6. show rows changed.

## Acceptance criteria
Deliver cleaned output and a reconciliation of rows before/after.

## Compatibility and failure modes
Avoid treating IDs as numbers; Excel dates are locale-sensitive.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
