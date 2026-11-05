---
name: financial-models
description: Use for budgets, revenue forecasts, cost scenarios, variances and cash projections.
---

# Create auditable budget and financial models

**Trigger:** Use for budgets, revenue forecasts, cost scenarios, variances and cash projections.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Separate assumptions, historical actuals and derived calculations.
2. use consistent periods and units.
3. link rather than hardcode.
4. implement scenario toggles.
5. test sign conventions, opening and closing balances and sensitivity.

## Acceptance criteria
Reconcile model to source records, expose inputs, and note unsupported assumptions.

## Compatibility and failure modes
Validate Excel 2010 formula support; add sources and dates to external assumptions.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
