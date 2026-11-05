---
name: recovery-repair
description: Use for files opening with repair dialogs, corrupt packages or lost charts.
---

# Investigate Excel repair and unreadable content warnings

**Trigger:** Use for files opening with repair dialogs, corrupt packages or lost charts.

## Required starting point
Obtain the source workbook or an explicit new-workbook specification; identify Excel version, expected output and must-preserve elements. Back up existing files. Check public-repository privacy before using examples.

## Implementation workflow
1. Keep the damaged original.
2. check ZIP CRC, XML parsing, required parts and relationships.
3. inspect workbook/sheet/chart/table references.
4. review Excel repair log.
5. rebuild only damaged elements in a copy.
6. compare features before and after.

## Acceptance criteria
Explain what was corrupt, what was repaired and what may have been lost.

## Compatibility and failure modes
A structurally valid XLSX can still trigger Excel-specific repair; test with target desktop Excel.

## Standard verification
Run `python scripts/audit_xlsx.py output.xlsx`, inspect source-versus-output structure where edits occur, independently test at least one normal and one exception scenario, and disclose untested Excel desktop behaviour. For tasks involving macros/ETL, apply extra environment-specific checks before claiming success.
