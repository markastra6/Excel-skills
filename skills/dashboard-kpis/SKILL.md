---
name: dashboard-kpis
description: Use when users need progress, productivity, remaining scope, quality and visual exceptions.
---

# Create connected Excel KPI dashboards

**Trigger:** Use when users need progress, productivity, remaining scope, quality and visual exceptions.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Define each KPI in writing.
2. connect metrics to canonical data ranges.
3. use unit-wise SUM(MAX(0, plan-actual)) style logic for backlog.
4. distinguish scope and productivity.
5. include drill-through or easy reference to source rows.

## Acceptance criteria
Cross-check dashboard against raw data and reconciliation totals; avoid duplicate data entry.

## Compatibility and failure modes
Keep display compact, with a fallback to the source table rather than unsupported slicers.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
