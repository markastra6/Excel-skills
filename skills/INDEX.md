# Excel Skills Catalog

**35 task-specific skills** are available as Markdown workflow instructions for assistants, developers, and Excel power users. They do not automatically install or enable Excel features. Read the corresponding `SKILL.md` before working.

| [Desktop Spreadsheet Workbench](desktop-spreadsheet-workbench/SKILL.md) | Native local Excel 2010+ tracker creation and updates |
| [Local App Integrations](local-app-integrations/SKILL.md) | GitHub export and safe multi-step workflows |

| Skill | When to use |
| --- | --- |
| [Analyze an existing workbook before edits](skills/inspect-workbook/SKILL.md) | Use for unknown Excel files, hidden data and layout discovery |
| [Create a new reusable Excel workbook](skills/create-workbook/SKILL.md) | Use when a new template, tracker, register or reporting workbook is requested |
| [Safely modify an existing workbook](skills/edit-preserve/SKILL.md) | Use for edits where formulas, metadata, staff data, formatting and charts must survive |
| [Build robust formulas and calculation logic](skills/formulas-engineering/SKILL.md) | Use for automatic totals, allocation rules, date math, quality metrics and reusable target controls |
| [Trace and correct formula failures](skills/formulas-audit/SKILL.md) | Use for #REF!, #VALUE!, wrong totals, hardcoded rates, circularity and unexpected cached results |
| [Reliable Excel lookups and joins](skills/lookup-reference/SKILL.md) | Use for matching IDs, chapter names, staff and source lists across sheets or workbooks |
| [Clean and standardize Excel data](skills/data-cleaning/SKILL.md) | Use for inconsistent text, dates, blanks, duplicates and import issues |
| [Set up editable tables and dropdowns](skills/tables-validation/SKILL.md) | Use when staff, units, categories or dates must be extendable |
| [Design a modern at-a-glance worksheet](skills/one-screen-layout/SKILL.md) | Use when users want everything readable without zooming out or scrolling |
| [Use colour to flag exceptions clearly](skills/conditional-formatting/SKILL.md) | Use for targets, performance, overdue work, status and quality warnings |
| [Build clear Excel charts and sparklines](skills/charts-visuals/SKILL.md) | Use for planned vs actual, chapter backlog, daily trends, and error distribution |
| [Create connected Excel KPI dashboards](skills/dashboard-kpis/SKILL.md) | Use when users need progress, productivity, remaining scope, quality and visual exceptions |
| [Summarize with PivotTables or formulas](skills/pivot-summary/SKILL.md) | Use for grouping by staff, unit, stage, date, correction category and project |
| [Design repeatable data import and transformations](skills/power-query-etl/SKILL.md) | Use for combining weekly files, folder data, CSV and externally refreshed sources |
| [Engineer safe Excel VBA automation](skills/vba-macros/SKILL.md) | Use for buttons, imports, user forms, workbook events or repeatable local actions |
| [Automate folders of Excel files on Windows](skills/batch-windows/SKILL.md) | Use for recurring project workbooks, batch audits, conversions and reports |
| [Make workbooks practical in Excel 2010](skills/excel-2010-compat/SKILL.md) | Use before releasing files to the user's local Excel 2010 environment |
| [Track production and QC per chapter or unit](skills/production-qc/SKILL.md) | Use for manuscript, proof, JA, composition, QC, correction rounds and deliveries |
| [Plan and measure staff workload](skills/team-capacity/SKILL.md) | Use for editable composer, QC and JA teams with daily targets and leave |
| [Build date-wise schedules and forecasting](skills/schedule-planning/SKILL.md) | Use for proof passes, chapter deliveries, project milestones and team capacity calendars |
| [Audit quality, classify errors and track RCA/CAPA](skills/quality-rca-capa/SKILL.md) | Use for PE/EA reviews, correction counts, QC findings and corrective actions |
| [Compare workbook versions or source records](skills/reconcile-compare/SKILL.md) | Use for before/after workbooks, source vs tracker, monthly revisions and data migrations |
| [Create auditable budget and financial models](skills/financial-models/SKILL.md) | Use for budgets, revenue forecasts, cost scenarios, variances and cash projections |
| [Build structured Excel registers](skills/inventory-register/SKILL.md) | Use for asset logs, project registers, issue trackers and attendance or stock entries |
| [Handle CSV/TSV and external data](skills/csv-import-export/SKILL.md) | Use for imports from tools, accounting apps, project systems and Excel-compatible export |
| [Prepare printable Excel reports and PDFs](skills/reporting-export/SKILL.md) | Use for handoff summaries, project reviews, billing reports and PDF printouts |
| [Diagnose slow or freezing Excel files](skills/workbook-performance/SKILL.md) | Use for Excel 2010 not responding, long recalculation or oversized workbooks |
| [Investigate Excel repair and unreadable content warnings](skills/recovery-repair/SKILL.md) | Use for files opening with repair dialogs, corrupt packages or lost charts |
| [Protect workbook data and safe sharing](skills/privacy-protection/SKILL.md) | Use for public GitHub publication, customer records, employee data and confidential spreadsheets |
| [Test and ship reliable Excel workbooks](skills/testing-release/SKILL.md) | Use before handing over a workbook or pushing a template to GitHub |
| [Turn a project workbook into a reusable master](skills/master-template/SKILL.md) | Use to convert a project-specific file into editable template for future jobs |
| [Upgrade legacy Excel without data loss](skills/migration-modernization/SKILL.md) | Use for .xls to .xlsx, Excel 2010 to newer Excel, or old formatting/templates |
| [Analyze larger tables with relationships and measures](skills/power-pivot-data-model/SKILL.md) | Use for Power Pivot, DAX and cross-table modelling when installed |

## Quick combinations
- **Existing project tracker:** inspect-workbook + edit-preserve + production-qc + team-capacity + charts-visuals + testing-release
- **One-screen, beautiful Excel:** one-screen-layout + dashboard-kpis + conditional-formatting + charts-visuals
- **Excel 2010 repair:** recovery-repair + excel-2010-compat + formulas-audit + testing-release
- **Reusable blank templates:** master-template + tables-validation + privacy-protection + testing-release
- **Batch Excel work:** batch-windows + inspect-workbook + reporting-export
- **Quality tracker:** quality-rca-capa + formulas-engineering + reconcile-compare + dashboard-kpis

## Reality check
Not every Excel command can be executed without Microsoft Excel or an automation engine. VBA, Power Query, Power Pivot and PDF export depend on local app capabilities; do not promise functionality that has not been tested. These skill files provide repeatable decisions, QA gates and implementation recipes.
