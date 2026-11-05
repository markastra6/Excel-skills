"""Offline test cases for the read-only workbook auditor."""
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from audit_xlsx import audit


class WorkbookAuditTests(unittest.TestCase):
    def make_book(self, directory, formula="1+1", missing_sheet=False, corrupt_xml=False):
        file = Path(directory) / "sample.xlsx"
        content = {
            "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
            "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="xl/workbook.xml"/></Relationships>',
            "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Sheet1" sheetId="1"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
            "xl/worksheets/sheet1.xml": '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>' + formula + '</f></c></row></sheetData></worksheet>',
        }
        if missing_sheet:
            del content["xl/worksheets/sheet1.xml"]
        if corrupt_xml:
            content["xl/workbook.xml"] = "<broken"
        with zipfile.ZipFile(file, "w") as z:
            for name, value in content.items():
                z.writestr(name, value)
        return file

    def test_valid(self):
        with tempfile.TemporaryDirectory() as d:
            report = audit(self.make_book(d))
            self.assertEqual(report["errors"], [])
            self.assertEqual(report["sheets"], 1)
            self.assertEqual(report["formulas"], 1)

    def test_broken_formula(self):
        with tempfile.TemporaryDirectory() as d:
            report = audit(self.make_book(d, formula="#REF!+1"))
            self.assertTrue(any("Broken formula" in e for e in report["errors"]))

    def test_missing_sheet_part(self):
        with tempfile.TemporaryDirectory() as d:
            report = audit(self.make_book(d, missing_sheet=True))
            self.assertTrue(any("Missing relationship target" in e for e in report["errors"]))

    def test_malformed_xml(self):
        with tempfile.TemporaryDirectory() as d:
            report = audit(self.make_book(d, corrupt_xml=True))
            self.assertTrue(any("Malformed XML" in e for e in report["errors"]))


if __name__ == "__main__":
    unittest.main()
