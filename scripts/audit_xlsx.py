#!/usr/bin/env python3
"""Read-only XLSX structure audit (standard library; no Excel dependency).

Does NOT recalculate formulas or certify Excel 2010 compatibility.
"""
import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REQUIRED = {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"}
ERROR_VALUES = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#N/A", "#NULL!"}


def audit(path):
    result = {"file": str(path), "sheets": 0, "charts": 0,
              "formulas": 0, "errors": [], "warnings": []}
    if not path.is_file():
        result["errors"].append("File not found")
        return result
    try:
        with zipfile.ZipFile(path) as archive:
            parts = set(archive.namelist())
            damaged = archive.testzip()
            if damaged:
                result["errors"].append("ZIP CRC failure: " + damaged)
            for name in sorted(REQUIRED - parts):
                result["errors"].append("Missing required part: " + name)
            roots = {}
            for name in sorted(parts):
                if name.endswith((".xml", ".rels")):
                    try:
                        roots[name] = ET.fromstring(archive.read(name))
                    except ET.ParseError as exc:
                        result["errors"].append("Malformed XML " + name + ": " + str(exc))
            wb = roots.get("xl/workbook.xml")
            if wb is not None:
                sheet_names = [e.get("name", "") for e in wb.findall(".//x:sheets/x:sheet", NS)]
                result["sheets"] = len(sheet_names)
                if len({n.casefold() for n in sheet_names}) != len(sheet_names):
                    result["errors"].append("Duplicate sheet names")
            result["charts"] = sum(bool(re.fullmatch(r"xl/charts/chart\\d+\\.xml", p)) for p in parts)
            if any("xl/charts/chartEx" in p for p in parts):
                result["warnings"].append("Modern chart extension: verify in Excel 2010")
            for name, root in roots.items():
                if name.endswith(".rels"):
                    directory = posixpath.dirname(name)
                    source_dir = directory[:-6] if directory.endswith("/_rels") else ""
                    for rel in root:
                        if rel.get("TargetMode") == "External":
                            continue
                        target = unquote((rel.get("Target") or "").split("#", 1)[0])
                        if not target:
                            result["errors"].append("Empty relationship: " + name)
                            continue
                        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(source_dir, target))
                        if resolved not in parts:
                            result["errors"].append("Missing relationship target: " + name + " -> " + resolved)
                if re.fullmatch(r"xl/worksheets/sheet\\d+\\.xml", name):
                    for f in root.findall(".//x:f", NS):
                        value = f.text or ""
                        result["formulas"] += 1
                        if "#REF!" in value:
                            result["errors"].append("Broken formula reference in " + name)
                        if "_xlfn." in value or "_xlws." in value:
                            result["warnings"].append("Potentially unsupported Excel 2010 formula in " + name)
                    for cell in root.findall(".//x:c", NS):
                        if cell.get("t") == "e":
                            v = cell.find("x:v", NS)
                            if v is not None and v.text in ERROR_VALUES:
                                result["errors"].append("Stored formula error " + v.text + " in " + name + "!" + str(cell.get("r")))
            result["warnings"] = sorted(set(result["warnings"]))
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        result["errors"].append("Cannot inspect workbook: " + str(exc))
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", type=Path, nargs="+")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    results = [audit(p) for p in args.files]
    if args.json:
        print(json.dumps(results, indent=2))
    else:
        for item in results:
            print(("FAIL" if item["errors"] else "PASS") + " " + item["file"]
                  + " | sheets=" + str(item["sheets"]) + " charts=" + str(item["charts"])
                  + " formulas=" + str(item["formulas"]))
            for field in ("errors", "warnings"):
                for msg in item[field]:
                    print("  " + field.upper() + ": " + msg)
    return 1 if any(r["errors"] for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
