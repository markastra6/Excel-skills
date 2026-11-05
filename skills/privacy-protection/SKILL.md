---
name: privacy-protection
description: Use for public GitHub publication, customer records, employee data and confidential spreadsheets.
---

# Protect workbook data and safe sharing

**Trigger:** Use for public GitHub publication, customer records, employee data and confidential spreadsheets.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Check visible and hidden sheets, comments, defined names, custom properties, cached pivot data, external links, connections and revision traces.
2. redact with an explicit data-minimization plan.
3. protect workbook where appropriate.
4. scan exports.

## Acceptance criteria
Only publish verified anonymized examples; use private storage for live project files.

## Compatibility and failure modes
Worksheet protection is not encryption; password-protected workbook handling requires explicit approval.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
