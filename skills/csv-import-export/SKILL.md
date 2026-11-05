---
name: csv-import-export
description: Use for imports from tools, accounting apps, project systems and Excel-compatible export.
---

# Handle CSV/TSV and external data

**Trigger:** Use for imports from tools, accounting apps, project systems and Excel-compatible export.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Detect encoding and delimiter.
2. preserve quoted newlines and leading-zero identifiers.
3. validate schema.
4. handle duplicate headers.
5. treat formulas beginning with =,+,-,@ as potential CSV injection in exported text.
6. compare source/destination counts.

## Acceptance criteria
Produce row-count reconciliation and a stable encoding/delimiter spec.

## Compatibility and failure modes
Excel 2010 opening CSV may coerce IDs, dates and encodings without import settings.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
