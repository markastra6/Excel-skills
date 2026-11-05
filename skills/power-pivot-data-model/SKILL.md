---
name: power-pivot-data-model
description: Use for Power Pivot, DAX and cross-table modelling when installed.
---

# Analyze larger tables with relationships and measures

**Trigger:** Use for Power Pivot, DAX and cross-table modelling when installed.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Define grain and primary keys.
2. model relationships explicitly.
3. write named measures with clear denominators.
4. compare totals to source.
5. test refresh behavior and supported model versions.

## Acceptance criteria
Document model fields, measures and refresh dependencies.

## Compatibility and failure modes
Power Pivot in Excel 2010 depends on edition/add-in availability and cannot be assumed.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
