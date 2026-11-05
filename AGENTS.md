# Excel Skills — agent guidance

This repository contains **opt-in workflow skills**, not automatically installed Excel plugins. Read [skills/INDEX.md](skills/INDEX.md), choose the smallest relevant skill set, and inspect the user's workbook **before** edits.

## Permanent Excel repair-warning prevention policy

**Never deliver an XLSX as Excel-2010-verified based solely on ZIP/XML tests or a third-party engine.**
All final workbooks must pass `python scripts/audit_xlsx.py --excel-2010-strict file.xlsx`; the release reviewer must also open, recalculate, save, close and reopen a COPY in desktop Excel 2010 with **no unreadable-content or repair prompt**. If desktop Excel is unavailable, state the limitation explicitly and do not claim guaranteed compatibility.

When a generated file contains literal cells typed as `t="str"`, ambiguous workbook XML MIME defaults, or unnecessary x14 conditional-formatting extensions, normalize the package using `scripts/repair_ooxml_compat.py` before delivery and rerun strict checks. Preserve formulas, cached values, existing data, styles, drawing/chart relationships and originals; compare source against output cell-by-cell. Never auto-run destructive Excel recovery on a client's original file. Use repair logs to identify any remaining cause.

This is a standing requirement, including Production/QC trackers, templates, charts and workbook redesigns.

## Non-negotiable rules
1. Preserve source files; write new outputs unless overwrite is specifically authorized.
2. Start from the actual source workbook—not an invented generic layout.
3. Maintain one source of truth for targets, employee allocations, dates and quality fields.
4. Distinguish manuscript pages, proof-page plans, completed pages, correction events and confirmed PE-affected pages.
5. Do not infer chapter attribution from aggregate day-wise output; expose differences.
6. Keep a modern, compact and readable at-a-glance view when requested. Use full labels, not unexplained abbreviations.
7. Target **Excel 2010** unless the user explicitly authorizes newer-only functionality.
8. Verify XML/ZIP package integrity **and** formulas/visual layout. State explicitly when desktop Excel verification is unavailable.
9. Never put client, staff or active project records in this public repository; anonymize hidden sheets, comments, names, cached data and metadata too.
10. Use native Excel formulas/charts/validations where they make sense. Python scripts in `scripts/` are read-only unless their documentation says otherwise.

## Tool selection

Use [desktop-spreadsheet-workbench](skills/desktop-spreadsheet-workbench/SKILL.md) for actual local Microsoft Excel actions. Use [local-app-integrations](skills/local-app-integrations/SKILL.md) for GitHub data exports or allowlisted local workflows. Never claim these replicate the proprietary @Spreadsheets or Composio services. Never embed credentials in repository files.

## Tools
- `python scripts/audit_xlsx.py workbook.xlsx`
- `python scripts/inventory_xlsx.py workbook.xlsx`
- `python scripts/compare_xlsx.py original.xlsx revised.xlsx`
- `python -m unittest discover -s tests -v`

For actual workbook creation/editing in supported environments, follow the installed spreadsheet engine's documented API; do not substitute a static screenshot for an editable workbook.
