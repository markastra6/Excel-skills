---
name: one-screen-layout
description: Use when users want everything readable without zooming out or scrolling.
---

# Design a modern at-a-glance worksheet

**Trigger:** Use when users want everything readable without zooming out or scrolling.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Keep key metrics in a top strip.
2. chapters on left and daily team figures on right when appropriate.
3. collapse redundant columns.
4. use precise full-length labels and two-line headers.
5. limit colours to meaning.
6. reserve space for compact charts.
7. inspect at normal display resolution.

## Acceptance criteria
Make input vs calculated cells obvious; provide a legible one-screen screenshot and the native workbook.

## Compatibility and failure modes
Do not sacrifice readable text for literal one-screen claims on unknown monitor sizes.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
