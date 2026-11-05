---
name: team-capacity
description: Use for editable composer, QC and JA teams with daily targets and leave.
---

# Plan and measure staff workload

**Trigger:** Use for editable composer, QC and JA teams with daily targets and leave.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Maintain one staff roster and role/rate rules.
2. distinguish available hours from leave and other-project work.
3. assign workload by date and unit.
4. calculate target based on working time and work type.
5. mark mixed SB/WS rate as unknown unless split is recorded.

## Acceptance criteria
Show planned capacity vs assigned load and flag overbooking, double allocation or absent rates.

## Compatibility and failure modes
Do not assume everyone works eight full hours or one work type each day.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
