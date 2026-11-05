# Excel Skills — agent guidance

This repository contains **opt-in workflow skills**, not automatically installed Excel plugins. Read [skills/INDEX.md](skills/INDEX.md), choose the smallest relevant skill set, and inspect the user's workbook **before** edits.

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
