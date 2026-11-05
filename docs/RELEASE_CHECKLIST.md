# Workbook release checklist

1. Back up source workbook; don't overwrite it during development.
2. Check input cells, target strip, dropdowns, formulas, status logic, conditional formatting, charts and print layout.
3. Change Comp, QC and JA targets and verify dependent results.
4. Test new staff, added chapters and dates, empty fields, zero values, mixed SB/WS assignments and output exceeding plan.
5. Check that unit-by-unit remaining pages and scope reconciliation behave correctly.
6. Confirm actual chart source ranges and readable on-screen layout.
7. Run `python scripts/audit_xlsx.py --excel-2010-strict file.xlsx` and **fail delivery** if ANY errors remain. If the file is generated with a nonconforming OOXML serializer, run `python scripts/repair_ooxml_compat.py source.xlsx output.xlsx` on a copy, then rerun the strict audit and verify every original formula/value and chart remains unchanged.
8. Mandatory desktop acceptance: Open a **copy** in Excel 2010 with no 'unreadable content' or 'recover contents' prompt, check the charts, validation lists, formulas, formatting and open/reopen behaviour. If unavailable, label 'Desktop Excel 2010 opening unverified'—never claim guaranteed compatibility.
9. Recalculate, inspect errors, save, close and reopen.
10. Verify no client, employee or project-specific records are included before publishing in this public repository.

The Python audit does not recalculate Excel formulas or guarantee successful opening in desktop Excel.

## Regression: 5 November 2026 unreadable-content incident

AL Decision Maths Production Studio generated a workbook with 397 literal text cells incorrectly marked `t="str"`, 40 empty cells with inappropriate string types, XML-wide workbook MIME defaults, redundant extended x14 data-bar metadata and a duplicated chart drawing name. Previous ZIP/XML audits passed because they verified only basic syntax and relationships. The exact warning trigger remains unconfirmed without the Microsoft Excel recovery log.

Regression protections: strict OOXML audit, conservative normalization, automated fixtures, chart/formula comparison and mandatory Excel 2010 desktop opening. **A structural pass never substitutes for an Excel open test.**
