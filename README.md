# Excel Skills

Reusable, Excel 2010-friendly workbooks and automation tools for book-production projects.

## What this repository is for

- **Production & QC control:** chapter/unit scope, JA, composition, QC, correction counts, quality, daily employee output and workload.
- **One-screen visibility:** compact, professional layouts with meaningful charts and colour-coded exceptions—not bulky multi-tab dashboards.
- **Single-point settings:** enter shared Comp, QC and JA targets once; dependent calculations update automatically.
- **Editable staffing:** support changes in composers, QC staff, chapters, work types and dates.
- **Release safety:** audit workbook package structure, chart relationships, stored formula errors and possible Excel 2010 compatibility risks.

## Repository structure

| Path | Purpose |
| --- | --- |
| `templates/` | **Anonymized, reusable** blank Excel templates |
| `scripts/` | Offline workbook validation and automation |
| `tests/` | Automated tests for scripts |
| `docs/` | Tracker design and release standards |
| `.github/workflows/` | Checks on GitHub pushes and pull requests |

## Excel skill catalog and tools

- [33 task-specific Excel skills](skills/INDEX.md)
- [Executable offline workbook tools](docs/TOOLS_AND_USAGE.md)
- [Guidance for coding assistants](AGENTS.md)
- [Production and QC template standards](docs/TRACKER_DESIGN.md)

## Check an Excel file

```bash
python scripts/audit_xlsx.py path/to/workbook.xlsx
python -m unittest discover -s tests -v
```

The audit checks ZIP/XML integrity, internal relationships, stored formula errors, and potentially incompatible functions. **It does not calculate formulas or prove a workbook will open cleanly in Excel 2010.** Test in desktop Excel before release.

## Privacy

**This repository is public.** Do not commit live customer manuscripts, employee names or output, client-specific project workbooks, or other confidential production records. Use sanitized blank templates only. Keep operational copies in approved internal storage.

See the [tracker design](docs/TRACKER_DESIGN.md) and [release checklist](docs/RELEASE_CHECKLIST.md).
