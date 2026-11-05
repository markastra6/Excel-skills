---
name: lookup-reference
description: Use for matching IDs, chapter names, staff and source lists across sheets or workbooks.
---

# Reliable Excel lookups and joins

**Trigger:** Use for matching IDs, chapter names, staff and source lists across sheets or workbooks.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Normalize keys without destroying originals.
2. choose exact match.
3. check duplicates and missing values.
4. use INDEX/MATCH or VLOOKUP(FALSE) with fixed reference ranges.
5. report unmatched and multiple matches.

## Acceptance criteria
Output match/unmatch counts and a reviewable exception list.

## Compatibility and failure modes
Excel 2010 does not support XLOOKUP; guard text-number and leading-zero IDs.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
