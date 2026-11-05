---
name: create-workbook
description: Use when a new template, tracker, register or reporting workbook is requested.
---

# Create a new reusable Excel workbook

**Trigger:** Use when a new template, tracker, register or reporting workbook is requested.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Define input schema, units and metrics.
2. plan clear section hierarchy.
3. centralize constants and target assumptions.
4. create validated input cells, copy-safe formulas, sensible number formats and at least one compact visual where useful.
5. set freeze panes and print area.

## Acceptance criteria
Deliver a real .xlsx plus a short input guide; test blank, normal and edge-case inputs and reopen exported file.

## Compatibility and failure modes
Use Excel 2010-supported formula functions by default; avoid dynamic arrays.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
