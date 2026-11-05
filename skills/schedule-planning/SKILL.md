---
name: schedule-planning
description: Use for proof passes, chapter deliveries, project milestones and team capacity calendars.
---

# Build date-wise schedules and forecasting

**Trigger:** Use for proof passes, chapter deliveries, project milestones and team capacity calendars.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Identify due dates and working calendar.
2. model dependencies and holidays.
3. convert page scope to hours using central rates.
4. allocate dates to available staff.
5. update projections from actual output.
6. flag late and blocked work.

## Acceptance criteria
Provide baseline vs current projection, bottlenecks and next deadlines.

## Compatibility and failure modes
Excel 2010 WORKDAY/NETWORKDAYS functions are available, but locale date parsing needs checks.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
