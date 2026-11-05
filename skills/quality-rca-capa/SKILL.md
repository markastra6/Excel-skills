---
name: quality-rca-capa
description: Use for PE/EA reviews, correction counts, QC findings and corrective actions.
---

# Audit quality, classify errors and track RCA/CAPA

**Trigger:** Use for PE/EA reviews, correction counts, QC findings and corrective actions.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Separate total corrections, PE events, unique PE-affected pages and EA.
2. define denominator for any quality percentage.
3. use classifications and employee response.
4. record root cause, containment, corrective and preventive action.
5. date ownership and closure.

## Acceptance criteria
Do not display a quality score without verified numerator and denominator; show pending instead.

## Compatibility and failure modes
One page may contain several errors; do not substitute correction count for affected pages.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
