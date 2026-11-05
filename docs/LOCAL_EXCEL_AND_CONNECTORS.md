# Local Excel + integration tools

## What's actually recreated

| Capability | Local tool | Limits |
| --- | --- | --- |
| Create/edit Excel workbooks | `desktop/ExcelSkills.ps1` via Excel COM | Needs Windows and desktop Excel 2010+ |
| Central Comp / QC / JA targets | NewTracker and SetTargets | Only ExcelSkills-generated tracker layouts |
| Charts and production overview | NewTracker | One chart and the workbook's native cells/formulas |
| Convert CSV to XLSX | ImportCsv | CSV opened as text to avoid unintended coercion |
| Excel PDF output | ExportPdf | Requires Excel and a PDF-capable installed version |
| Formula recalculation | Recalculate | Needs Excel; existing external links are not refreshed |
| GitHub issues to CSV | `integrations/github_issues.py` | GitHub only, token optional for public repositories |
| Multi-step automation | `integrations/workflow.py` | Allowlisted operations; not hosted Composio |
| Scheduled GitHub export | GitHub Actions | Runs when enabled by the repository owner |

The repository does **not** contain proprietary @Spreadsheets or Composio source, cloud accounts, installation tokens, or a clone of their services.

## Start using the desktop Excel tool

On **Windows with desktop Excel installed**:

1. Download/clone this repository, unzip it to a local folder.
2. Double-click `desktop/ExcelSkills.bat`.
3. Choose Create tracker, Update targets, Import CSV, PDF export, Recalculate, or Inspect.
4. Select new output locations; the script does not overwrite an existing file.
5. Check the output in Excel 2010 and validate with `python scripts/audit_xlsx.py new-file.xlsx`.

### PowerShell examples

```powershell
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action NewTracker -Destination "C:\Work\QC_Tracker.xlsx" -Project "New Project"
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action SetTargets -Source "C:\Work\QC_Tracker.xlsx" -Destination "C:\Work\QC_Tracker_v2.xlsx" -CompSB 30 -CompWS 20 -QCSB 40 -QCWS 25 -JA 1 -Hours 8
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action ImportCsv -Source "C:\Work\issues.csv" -Destination "C:\Work\issues.xlsx"
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action Inspect -Source "C:\Work\QC_Tracker.xlsx"
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action Recalculate -Source "C:\Work\QC_Tracker.xlsx" -Destination "C:\Work\QC_Tracker_Calculated.xlsx"
powershell.exe -NoProfile -File .\desktop\ExcelSkills.ps1 -Action ExportPdf -Source "C:\Work\QC_Tracker.xlsx" -Destination "C:\Work\QC_Tracker.pdf"
```

**Do not store original confidential project spreadsheets in this public GitHub repository.**

## GitHub integration

```bash
# Read-only plan, no network or file writes:
python integrations/workflow.py examples/github_issues_workflow.json

# Execute:
python integrations/workflow.py examples/github_issues_workflow.json --apply

# Direct export:
python integrations/github_issues.py --repo markastra6/Excel-skills --output reports/github_issues.csv --apply
```

The CSV is UTF-8 with BOM, so it opens more reliably in Excel 2010. Long/quoted cells and Tamil text are preserved by Python's CSV writer. Dangerous formula-like values are prefixed as literal text before export.

## Private repository access

If necessary, create a GitHub token with **read-only Issues permission** scoped to the relevant repository. Set it only in a secure environment:

```powershell
$env:GITHUB_TOKEN = Read-Host 'GitHub token'
```

Do not store tokens in workflow JSON files, shell scripts, code, Excel cells, or the public repository. For GitHub Actions, use the built-in `GITHUB_TOKEN` with least privilege.

## Validation

```bash
python -m unittest discover -s tests -v
python scripts/audit_xlsx.py path/to/file.xlsx
python scripts/inventory_xlsx.py path/to/file.xlsx
```

GitHub CI validates Python syntax, runs tests, parses the PowerShell script, and structurally audits any public blank templates. The Windows COM automation and PDF export cannot be runtime-tested on Linux GitHub runners; verify those on a Windows computer running Excel 2010.

## Extension policy

New connectors should validate input IDs and destination paths, use least-privilege access, avoid printing credentials, use timeouts/retries, and be dry-run capable. Do not create arbitrary generic HTTP execution from JSON workflow files.
