"""Smoke tests for XLSX inventory, comparison and privacy scanner."""
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from xlsx_core import workbook_inventory, cell_map
from compare_xlsx import compare
from privacy_scan_xlsx import scan


def make_book(path, value="Alpha", hidden=False, formula="A1&\"!\""):
    sheet_xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:B2"/><sheetData>'
        '<row r="1"><c r="A1" t="inlineStr"><is><t>' + value + '</t></is></c>'
        '<c r="B1"><f>' + formula.replace("&","&amp;") + '</f><v/></c></row>'
        '</sheetData><mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells>'
        '<dataValidations count="1"><dataValidation sqref="A1" type="list"/></dataValidations>'
        '</worksheet>'
    )
    state = ' state="hidden"' if hidden else ''
    contents = {
        "[Content_Types].xml": "<Types/>",
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1"' + state + ' r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": sheet_xml,
    }
    with zipfile.ZipFile(path, "w") as z:
        for name, content in contents.items():
            z.writestr(name, content)
    return path


class ExcelToolkitTests(unittest.TestCase):
    def test_inventory(self):
        with tempfile.TemporaryDirectory() as d:
            p = make_book(Path(d) / "a.xlsx")
            inv = workbook_inventory(p)
            self.assertEqual(inv["sheets"][0]["formulas"], 1)
            self.assertEqual(inv["sheets"][0]["merged_ranges"], 1)
            self.assertEqual(inv["sheets"][0]["validation_rules"], 1)

    def test_compare(self):
        with tempfile.TemporaryDirectory() as d:
            a = make_book(Path(d) / "a.xlsx", value="Alpha")
            b = make_book(Path(d) / "b.xlsx", value="Beta")
            result = compare(a, b)
            self.assertEqual(result["changed_cells"], 1)
            self.assertEqual(result["added_cells"], 0)

    def test_privacy(self):
        with tempfile.TemporaryDirectory() as d:
            p = make_book(Path(d) / "a.xlsx", value="private@example.com", hidden=True)
            report = scan(p)
            kinds = [x["kind"] for x in report["findings"]]
            self.assertIn("email", kinds)
            self.assertIn("hidden_sheet", kinds)
            self.assertFalse(report["clear"])

    def test_safe_fixture(self):
        with tempfile.TemporaryDirectory() as d:
            p = make_book(Path(d) / "a.xlsx")
            self.assertTrue(scan(p)["clear"])
            self.assertEqual(len(cell_map(p)), 2)


if __name__ == "__main__":
    unittest.main()
