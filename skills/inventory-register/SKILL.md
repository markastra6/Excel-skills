---
name: inventory-register
description: Use for asset logs, project registers, issue trackers and attendance or stock entries.
---

# Build structured Excel registers

**Trigger:** Use for asset logs, project registers, issue trackers and attendance or stock entries.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Define unique IDs, fields, status codes, dates, audit trail and data validations.
2. calculate aging and overdue measures.
3. add duplicates/invalids summary.
4. ensure new items inherit formulas.

## Acceptance criteria
Test inserts, deletes, filtering, duplicates and exports.

## Compatibility and failure modes
No reliable audit trail from formulas alone—consider an external append-only log or VBA if needed.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
