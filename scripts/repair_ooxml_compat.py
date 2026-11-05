#!/usr/bin/env python3
"""Repair known OOXML compatibility defects in generator-produced XLSX files.

- Convert ordinary text erroneously saved as formula-string cells into inline text.
- Retain basic Excel 2007+ data bars while dropping redundant Excel 2010 x14 extensions.
- Provide explicit workbook ContentType and unique drawing object names.
- Never overwrite source or destination files.

This cannot guarantee Excel desktop will open a workbook without repair; desktop
Excel validation remains a release requirement. Standard library only.
"""
import argparse
import os
import sys
import tempfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
XML = "http://www.w3.org/XML/1998/namespace"
FALLBACK_ID = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
X14_BLOCK_ID = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"


def tag(ns, local):
    return "{" + ns + "}" + local


def fix_worksheet(data, count):
    root = ET.fromstring(data)
    for cell in root.iter(tag(MAIN, "c")):
        if cell.get("t") != "str" or cell.find(tag(MAIN, "f")) is not None:
            continue
        value = cell.find(tag(MAIN, "v"))
        if value is None:
            cell.attrib.pop("t", None)
            count["blank_string_type_cleared"] += 1
            continue
        literal = value.text or ""
        pos = list(cell).index(value)
        cell.remove(value)
        cell.set("t", "inlineStr")
        inline = ET.Element(tag(MAIN, "is"))
        text_node = ET.SubElement(inline, tag(MAIN, "t"))
        if literal.startswith((" ", "\t", "\n", "\r")) or literal.endswith((" ", "\t", "\n", "\r")):
            text_node.set(tag(XML, "space"), "preserve")
        text_node.text = literal
        cell.insert(pos, inline)
        count["literal_text_normalized"] += 1
    for rule in root.iter(tag(MAIN, "cfRule")):
        extension_list = rule.find(tag(MAIN, "extLst"))
        if extension_list is None:
            continue
        for ext in list(extension_list):
            if ext.get("uri", "").upper() == FALLBACK_ID:
                extension_list.remove(ext)
                count["redundant_cf_extension_removed"] += 1
        if len(extension_list) == 0:
            rule.remove(extension_list)
    for extension_list in list(root):
        if extension_list.tag != tag(MAIN, "extLst"):
            continue
        for ext in list(extension_list):
            if ext.get("uri", "").upper() == X14_BLOCK_ID:
                extension_list.remove(ext)
                count["x14_databar_metadata_removed"] += 1
        if len(extension_list) == 0:
            root.remove(extension_list)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def fix_drawing(data, count):
    root = ET.fromstring(data)
    names = set()
    for node in root.iter(tag(DRAW, "cNvPr")):
        name = node.get("name", "")
        if not name or name in names:
            candidate = "Drawing object " + node.get("id", "")
            counter = 1
            while candidate in names:
                counter += 1
                candidate = "Drawing object " + node.get("id", "") + "-" + str(counter)
            node.set("name", candidate)
            count["duplicate_drawing_names_corrected"] += 1
            name = candidate
        names.add(name)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def fix_content_types(data, count):
    root = ET.fromstring(data)
    default = next((e for e in root
                    if e.tag == tag(TYPES, "Default") and e.get("Extension") == "xml"), None)
    override = next((e for e in root
                     if e.tag == tag(TYPES, "Override") and e.get("PartName") == "/xl/workbook.xml"), None)
    if default is not None and default.get("ContentType") != "application/xml":
        default.set("ContentType", "application/xml")
        count["xml_content_type_corrected"] += 1
    if override is None:
        override = ET.SubElement(root, tag(TYPES, "Override"))
        override.set("PartName", "/xl/workbook.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        count["workbook_override_added"] += 1
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def repair(source, output):
    source, output = Path(source), Path(output)
    if source.suffix.lower() != ".xlsx" or output.suffix.lower() != ".xlsx":
        raise ValueError("Only XLSX files are supported")
    if not source.is_file():
        raise FileNotFoundError(source)
    if source.resolve() == output.resolve():
        raise ValueError("Output must differ from source")
    if output.exists():
        raise FileExistsError(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".excelcompat-", suffix=".xlsx", dir=output.parent)
    os.close(fd)
    counts = Counter()
    try:
        with ZipFile(source) as src, ZipFile(tmp, "w") as dst:
            if src.testzip() is not None:
                raise ValueError("Source ZIP has a CRC error")
            for info in src.infolist():
                value = src.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    value = fix_content_types(value, counts)
                elif info.filename.startswith("xl/worksheets/sheet") and info.filename.endswith(".xml"):
                    value = fix_worksheet(value, counts)
                elif info.filename.startswith("xl/drawings/drawing") and info.filename.endswith(".xml"):
                    value = fix_drawing(value, counts)
                dst.writestr(info, value)
        with ZipFile(tmp) as verify:
            if verify.testzip() is not None:
                raise ValueError("Output ZIP has a CRC error")
        os.replace(tmp, output)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return dict(counts)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        report = repair(args.source, args.output)
    except (OSError, ValueError, ET.ParseError) as exc:
        print("FAILED: " + str(exc), file=sys.stderr)
        return 1
    print("Created: " + str(args.output))
    for name, number in sorted(report.items()):
        print("  " + name + ": " + str(number))
    return 0


if __name__ == "__main__":
    sys.exit(main())
