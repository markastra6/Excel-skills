---
name: power-query-etl
description: Use for combining weekly files, folder data, CSV and externally refreshed sources.
---

# Design repeatable data import and transformations

**Trigger:** Use for combining weekly files, folder data, CSV and externally refreshed sources.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Specify source locations, schema, refresh trigger, join keys and transformation log.
2. use Power Query only on compatible Excel versions.
3. provide a fallback CSV/Excel import for older installations.

## Acceptance criteria
Show row counts, source freshness and refresh/error handling.

## Compatibility and failure modes
Power Query support in Excel 2010 requires a compatible add-in and may not be installed.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
