---
name: reporting-export
description: Use for handoff summaries, project reviews, billing reports and PDF printouts.
---

# Prepare printable Excel reports and PDFs

**Trigger:** Use for handoff summaries, project reviews, billing reports and PDF printouts.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Define audience and period.
2. use only validated metrics.
3. repeat header rows, set print area, orientation, margins and page numbers.
4. validate charts and no clipped tables.
5. retain editable .xlsx source.

## Acceptance criteria
Deliver both editable workbook and PDF only when requested.

## Compatibility and failure modes
Printing fidelity varies by installed fonts, printer drivers and Excel version.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
