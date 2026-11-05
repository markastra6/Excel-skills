# Workbook release checklist

1. Back up source workbook; don't overwrite it during development.
2. Check input cells, target strip, dropdowns, formulas, status logic, conditional formatting, charts and print layout.
3. Change Comp, QC and JA targets and verify dependent results.
4. Test new staff, added chapters and dates, empty fields, zero values, mixed SB/WS assignments and output exceeding plan.
5. Check that unit-by-unit remaining pages and scope reconciliation behave correctly.
6. Confirm actual chart source ranges and readable on-screen layout.
7. Run `python scripts/audit_xlsx.py file.xlsx` and fix errors.
8. Open a **copy** in Excel 2010. Check for repair prompts, unexpected fill colours and missing chart elements.
9. Recalculate, inspect errors, save, close and reopen.
10. Verify no client, employee or project-specific records are included before publishing in this public repository.

The Python audit does not recalculate Excel formulas or guarantee successful opening in desktop Excel.
