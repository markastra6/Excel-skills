---
name: desktop-spreadsheet-workbench
description: Create, edit, import, recalculate, inspect and export native Excel workbooks locally with Microsoft Excel 2010 or newer on Windows.
---

# Desktop Spreadsheet Workbench

This is an **open, locally executable alternative workflow** to a connected spreadsheet editor. It is not the proprietary @Spreadsheets plugin.

## When to use
- Creating a compact, one-sheet Production/QC tracker with centrally editable rates.
- Updating shared Comp, QC and JA targets without manually repairing formulas.
- Importing text-safe CSV; exporting a PDF; recalculating or inspecting workbook files.

## Commands
On a Windows PC with desktop Excel installed:

```bat
desktop\ExcelSkills.bat
```

Or use PowerShell noninteractively:

```powershell
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action NewTracker -Destination "C:\Reports\Blank_Tracker.xlsx" -Project "New Project"
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action SetTargets -Source "C:\Reports\Blank_Tracker.xlsx" -Destination "C:\Reports\Tracker_Revised.xlsx" -CompSB 35 -CompWS 20 -QCSB 45 -QCWS 25 -JA 1 -Hours 8
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action ImportCsv -Source "C:\Reports\Issues.csv" -Destination "C:\Reports\Issues.xlsx"
```

## Requirements and boundaries
- Uses Microsoft Excel COM automation; it cannot run on Linux/macOS or without desktop Excel.
- Tested in CI for script parsing only. Native Excel 2010 execution must be tested on Windows.
- The generated starter tracker is a **new** format; `SetTargets` only works on ExcelSkills-generated workbooks. It does not automatically edit arbitrary legacy trackers.
- Workbook processing must leave the original source unchanged; outputs must use new filenames.
- CSV import preserves imported text including leading-zero IDs and avoids automatic Excel date coercion.
- Verify the output using `scripts/audit_xlsx.py` and desktop Excel.
