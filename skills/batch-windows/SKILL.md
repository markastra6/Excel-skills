---
name: batch-windows
description: Use for recurring project workbooks, batch audits, conversions and reports.
---

# Automate folders of Excel files on Windows

**Trigger:** Use for recurring project workbooks, batch audits, conversions and reports.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Prefer a .bat entrypoint prompting for source and output paths.
2. support Unicode and quoted paths, UNC shares, long paths, retries and logs.
3. do not overwrite originals without explicit consent.
4. provide dry-run and clear exit codes.

## Acceptance criteria
Test on spaces in names, network shares, permission failures and interruptions.

## Compatibility and failure modes
Avoid assuming Excel COM available for unattended servers; warn about Excel desktop session requirements.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
