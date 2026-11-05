#!/usr/bin/env python3
"""Summarize workbook sheets, formulas, charts, validations and hidden objects."""
import argparse
import json
import sys
from pathlib import Path
from xlsx_core import workbook_inventory


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)
    reports = []
    failures = 0
    for path in args.files:
        try:
            reports.append(workbook_inventory(path))
        except Exception as exc:
            failures += 1
            reports.append({"path": str(path), "error": type(exc).__name__ + ": " + str(exc)})
    if args.json:
        print(json.dumps(reports, indent=2))
    else:
        for report in reports:
            print("\n" + report["path"])
            if "error" in report:
                print("  ERROR: " + report["error"])
                continue
            print("  Charts: {charts}; tables: {tables}; images: {images}; macros: {macros}; external links: {external_links}".format(**report))
            for sheet in report["sheets"]:
                print("  - {name} [{state}]  {dimension}  cells={cells} formulas={formulas} merged={merged_ranges} validations={validation_rules} conditional_blocks={conditional_format_blocks}".format(**sheet))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
