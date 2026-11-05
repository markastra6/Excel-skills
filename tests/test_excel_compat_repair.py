"""Regression tests for Excel repair warning prevention (offline, stdlib only)."""
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from audit_xlsx import audit
from repair_ooxml_compat import repair

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"


def fixture(path):
    items = {
        "[Content_Types].xml":
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml":
            '<workbook xmlns="' + S + '"><sheets><sheet name="CONTROL" sheetId="1"/></sheets></workbook>',
        "xl/worksheets/sheet1.xml":
            '<worksheet xmlns="' + S + '"><sheetData><row r="1">'
            '<c r="A1" t="str"><v>Actual text</v></c>'
            '<c r="B1" t="str"><f>IF(1=1,"yes","no")</f><v>yes</v></c>'
            '<c r="C1" t="str"/>'
            '</row></sheetData></worksheet>',
        "xl/drawings/drawing1.xml":
            '<xdr:wsDr xmlns:xdr="' + D + '">'
            '<xdr:cNvPr id="1" name="Chart"/>'
            '<xdr:cNvPr id="2" name="Chart"/>'
            '</xdr:wsDr>',
    }
    with zipfile.ZipFile(path, "w") as output:
        for name, content in items.items():
            output.writestr(name, content)
    return path


class Excel2010RepairRegression(unittest.TestCase):
    def test_bad_generator_package_rejected_by_strict_audit(self):
        with tempfile.TemporaryDirectory() as d:
            file = fixture(Path(d) / "bad.xlsx")
            report = audit(file, excel_2010_strict=True)
            self.assertTrue(any("formula-string" in e for e in report["errors"]))
            self.assertTrue(any("drawing object" in e for e in report["errors"]))
            self.assertTrue(any("content type" in e for e in report["errors"]))

    def test_compatibility_repair_keeps_formulas_and_text(self):
        with tempfile.TemporaryDirectory() as d:
            file = fixture(Path(d) / "before.xlsx")
            output = Path(d) / "after.xlsx"
            counts = repair(file, output)
            self.assertEqual(counts["literal_text_normalized"], 1)
            self.assertEqual(counts["blank_string_type_cleared"], 1)
            self.assertEqual(counts["duplicate_drawing_names_corrected"], 1)
            report = audit(output, excel_2010_strict=True)
            self.assertEqual(report["errors"], [])
            with zipfile.ZipFile(output) as z:
                root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
                cells = {c.get("r"): c for c in root.iter("{" + S + "}c")}
                self.assertEqual(cells["A1"].get("t"), "inlineStr")
                self.assertEqual(cells["A1"].find(".//{" + S + "}t").text, "Actual text")
                self.assertEqual(cells["B1"].find("{" + S + "}f").text, 'IF(1=1,"yes","no")')
                self.assertIsNone(cells["C1"].get("t"))
            self.assertEqual(file.read_bytes(), file.read_bytes())


if __name__ == "__main__":
    unittest.main()
