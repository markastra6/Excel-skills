---
name: master-template
description: Use to convert a project-specific file into editable template for future jobs.
---

# Turn a project workbook into a reusable master

**Trigger:** Use to convert a project-specific file into editable template for future jobs.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Keep professional design.
2. preserve formulas and validations.
3. centralize targets and project info.
4. remove names, IDs and operational records including hidden objects.
5. test capacity with added units and staff.
6. document how to start fresh.

## Acceptance criteria
Deliver a genuinely blank master and a populated example only when it is safe to share.

## Compatibility and failure modes
Blanking visible cells alone does not anonymize embedded metadata or historical caches.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
