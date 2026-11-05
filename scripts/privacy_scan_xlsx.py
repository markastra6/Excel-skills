#!/usr/bin/env python3
"""Read-only privacy preflight for XLSX publication; never prints matched private values.

Heuristic: an empty report does NOT guarantee anonymity. Review hidden data and
objects manually. Does not remove content.
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xlsx_core import cell_map, workbook_inventory

EMAIL = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")
WINDOWS_USER = re.compile(r"(?:[A-Za-z]:\\Users\\|\\\\[^\\\s]+\\[^\\\s]+)")
PHONE = re.compile(r"(?<!\d)(?:\+91[- ]?)?[6-9][0-9]{9}(?!\d)")


def scan(path, terms=()):
    findings = []
    inventory = workbook_inventory(path)
    for sh in inventory["sheets"]:
        if sh["state"] != "visible":
            findings.append({"kind": "hidden_sheet", "location": sh["name"]})
    for label in ("external_links", "macros"):
        if inventory[label]:
            findings.append({"kind": label, "location": "workbook"})
    if inventory["images"]:
        findings.append({"kind": "embedded_images_need_review", "location": "workbook"})
    if inventory["defined_names"]:
        findings.append({"kind": "defined_names_need_review", "location": "workbook"})
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith(("xl/comments", "xl/threadedComments", "xl/pivotCache")):
                findings.append({"kind": "embedded_metadata_need_review", "location": name})
            if name in ("docProps/core.xml", "docProps/custom.xml"):
                findings.append({"kind": "document_properties_need_review", "location": name})
    patterns = [(EMAIL, "email"), (WINDOWS_USER, "local_or_network_path"), (PHONE, "phone")]
    for (sheet, ref), datum in cell_map(path).items():
        if datum[0] != "value" or datum[2] is None:
            continue
        value = str(datum[2])
        for pattern, kind in patterns:
            if pattern.search(value):
                findings.append({"kind": kind, "location": sheet + "!" + ref})
        for term in terms:
            if term and term.casefold() in value.casefold():
                findings.append({"kind": "restricted_term", "location": sheet + "!" + ref})
    return {"file": str(path), "findings": findings, "clear": not findings}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--term", action="append", default=[], help="Confidential term (never echoed)")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)
    results = []
    for path in args.files:
        try:
            results.append(scan(path, args.term))
        except Exception as exc:
            results.append({"file": str(path), "clear": False,
                            "findings": [{"kind": "inspection_error", "location": type(exc).__name__}]})
    if args.json:
        print(json.dumps(results, indent=2))
    else:
        for r in results:
            print(("REVIEW" if not r["clear"] else "NO HEURISTIC FLAGS") + ": " + r["file"])
            for f in r["findings"]:
                print("  " + f["kind"] + " at " + f["location"])
    return int(any(not r["clear"] for r in results))


if __name__ == "__main__":
    sys.exit(main())
