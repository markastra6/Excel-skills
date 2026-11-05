# Ready-to-run Excel command-line tools

No pip packages are required for these read-only tools. Use **Python 3.9+**. They accept paths containing spaces and support .xlsx files only.

| Command | Action | Modifies file? |
| --- | --- | --- |
| `python scripts/audit_xlsx.py "My File.xlsx"` | Verify ZIP, XML, relationships, stored formula errors | No |
| `python scripts/inventory_xlsx.py "My File.xlsx"` | Count sheets, formulas, merges, dropdowns, charts and hidden data | No |
| `python scripts/compare_xlsx.py "Before.xlsx" "After.xlsx"` | Report added, removed, changed literal/formula cells and structure | No |
| `python scripts/privacy_scan_xlsx.py "Blank Template.xlsx" --term "Client Name"` | Highlight cells requiring a privacy review without printing private values | No |
| `scripts/batch_audit_windows.bat` | Interactively audit every XLSX under a Windows source folder | No |

All Python tools support `--json`. The Windows batch file works with quoted paths, spaces and UNC locations. It prompts for a source folder and an output log file. It does not modify Excel workbooks.

## Important limitations

- Structural checking is **not** Excel recalculation or a desktop Excel 2010 open/repair test.
- Formula cached results are ignored by `compare_xlsx.py` to avoid treating stale caches as meaningful changes.
- `privacy_scan_xlsx.py` is a conservative indicator, not a certificate of anonymization. Review metadata, Excel cached data, images and confidential project records manually.
- Encrypted workbooks, legacy `.xls` and binary `.xlsb` are not supported by these XML-based tools.
- VBA, Power Query, Power Pivot, PivotTables and macros require a compatible installed Excel environment for their runtime actions.
