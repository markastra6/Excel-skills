"""Read-only XLSX parsing helpers (standard library, no Excel needed).

Supported: ordinary OOXML .xlsx, including shared/inline strings.
Not supported: encrypted workbooks, binary .xlsb or legacy .xls.
"""
import posixpath
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
X = lambda tag: "{" + MAIN + "}" + tag


def load_xml(archive, name):
    return ET.fromstring(archive.read(name))


def resolve_part(source, target):
    if not target:
        raise ValueError("Empty relationship target")
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(
        posixpath.join(posixpath.dirname(source), target)
    )


def sheets(archive):
    workbook = load_xml(archive, "xl/workbook.xml")
    relations = load_xml(archive, "xl/_rels/workbook.xml.rels")
    lookup = {r.get("Id"): resolve_part("xl/workbook.xml", r.get("Target"))
              for r in relations if r.get("Type", "").endswith("/worksheet")}
    result = []
    for sheet in workbook.findall("./" + X("sheets") + "/" + X("sheet")):
        rid = sheet.get("{" + OFFICE_REL + "}id")
        path = lookup.get(rid)
        if path is not None:
            result.append({"name": sheet.get("name", ""), "path": path,
                           "state": sheet.get("state", "visible")})
    return result, workbook


def shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = load_xml(archive, "xl/sharedStrings.xml")
    return ["".join(t.text or "" for t in item.iter(X("t")))
            for item in root.findall(X("si"))]


def decoded_cell(cell, strings):
    address = cell.get("r", "")
    kind = cell.get("t", "")
    f = cell.find(X("f"))
    v = cell.find(X("v"))
    value = v.text if v is not None else None
    if kind == "s" and value is not None:
        try:
            value = strings[int(value)]
        except (ValueError, IndexError):
            value = "<invalid shared-string reference>"
    elif kind == "inlineStr":
        inline = cell.find(X("is"))
        value = "".join(t.text or "" for t in inline.iter(X("t"))) if inline is not None else ""
    return address, value, f.text if f is not None else None, kind


def iter_sheet_cells(archive, sheet, strings=None):
    if strings is None:
        strings = shared_strings(archive)
    root = load_xml(archive, sheet["path"])
    for cell in root.iter(X("c")):
        entry = decoded_cell(cell, strings)
        if entry[1] is not None or entry[2] is not None:
            yield entry


def workbook_inventory(path):
    result = {"path": str(path), "sheets": [], "charts": 0, "tables": 0,
              "images": 0, "macros": False, "external_links": 0,
              "defined_names": 0}
    with zipfile.ZipFile(path) as archive:
        parts = archive.namelist()
        strings = shared_strings(archive)
        mapping, book = sheets(archive)
        result["charts"] = len([p for p in parts if p.startswith("xl/charts/chart") and p.endswith(".xml")])
        result["tables"] = len([p for p in parts if p.startswith("xl/tables/table") and p.endswith(".xml")])
        result["images"] = len([p for p in parts if p.startswith("xl/media/") and not p.endswith("/")])
        result["macros"] = "xl/vbaProject.bin" in parts
        result["external_links"] = len([p for p in parts if p.startswith("xl/externalLinks/externalLink") and p.endswith(".xml")])
        result["defined_names"] = len(book.findall(".//" + X("definedName")))
        for sheet in mapping:
            root = load_xml(archive, sheet["path"])
            cells = list(iter_sheet_cells(archive, sheet, strings))
            result["sheets"].append({
                "name": sheet["name"], "state": sheet["state"],
                "cells": len(cells),
                "formulas": sum(f is not None for _, _, f, _ in cells),
                "dimension": (root.find(X("dimension")).get("ref", "") if root.find(X("dimension")) is not None else ""),
                "merged_ranges": len(list(root.iter(X("mergeCell")))),
                "validation_rules": len(list(root.iter(X("dataValidation")))),
                "conditional_format_blocks": len(list(root.iter(X("conditionalFormatting")))),
                "protected": root.find(X("sheetProtection")) is not None,
            })
    return result


def cell_map(path):
    with zipfile.ZipFile(path) as archive:
        strings = shared_strings(archive)
        output = {}
        for sheet in sheets(archive)[0]:
            for addr, value, formula, kind in iter_sheet_cells(archive, sheet, strings):
                # The cached value of a formula is not reliable without Excel recalculation.
                output[(sheet["name"], addr)] = (
                    ("formula", formula) if formula is not None else ("value", kind, value)
                )
    return output
