import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import generate_quote as quote


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def cell_value(root, ref):
    for cell in root.iter(f"{NS_MAIN}c"):
        if cell.attrib.get("r") != ref:
            continue
        inline = cell.find(f"{NS_MAIN}is")
        if inline is not None:
            return "".join(t.text or "" for t in inline.iter(f"{NS_MAIN}t"))
        value = cell.find(f"{NS_MAIN}v")
        return value.text if value is not None else ""
    return ""


def cell_style(root, ref):
    for cell in root.iter(f"{NS_MAIN}c"):
        if cell.attrib.get("r") == ref:
            return int(cell.attrib["s"])
    raise AssertionError(f"Cell {ref} was not written")


def border_for_style(styles_root, style_id):
    cell_xfs = styles_root.find(f"{NS_MAIN}cellXfs")
    borders = styles_root.find(f"{NS_MAIN}borders")
    border_id = int(cell_xfs[style_id].attrib.get("borderId", "0"))
    return borders[border_id]


def generate_layout_workbook():
    brief = {
        "company_identity": "Koncept Image",
        "quote_date": "2026-06-04",
        "project_number": "KI-TEST-001",
        "client": {
            "name": "Sample Client",
            "attention": "Alex Tan",
        },
        "project": {
            "title": "Sample Project",
        },
        "line_items": [],
        "payment_terms": [],
        "signature": {
            "koncept_signatory": "Francies Cheng",
            "koncept_title": "Director",
        },
    }
    price = quote.PriceRow(
        row_number=1,
        section="Graphics",
        description="m2 of vinyl printed graphics",
        unit_hint="sqm",
        cost=100,
        gst_multiplier=1.09,
        markup=1,
        remark="",
    )
    line = quote.QuoteLine(
        section="Graphics",
        quantity=24,
        unit="m length",
        description="Vinyl printed graphics for a long booth fascia",
        pricing_keyword="vinyl printed graphics",
        display_price="",
        matched_price=price,
        amount=2400,
        match_status="matched",
        match_candidates=[],
    )
    tmp = tempfile.TemporaryDirectory()
    path = Path(tmp.name) / "quotation.xlsx"
    quote.write_quote_layout_xlsx(ROOT / "references" / "quotation-layout.xlsx", path, brief, [line])
    return tmp, path


class GenerateQuoteRowsTest(unittest.TestCase):
    def test_build_quote_rows_preserves_display_price_text(self):
        brief = {
            "company_identity": "Koncept Image",
            "quote_date": "2026-06-04",
            "client": {
                "name": "Sample Client",
                "attention": "Alex Tan",
            },
            "project": {
                "title": "Sample Project",
            },
            "line_items": [],
            "payment_terms": [],
        }
        line = quote.QuoteLine(
            section="Furniture Rental",
            quantity=1,
            unit="lot",
            description="Artificial green wall panels and potted plants",
            pricing_keyword="",
            display_price="Included",
            matched_price=None,
            amount=None,
            match_status="manual-display",
            match_candidates=[],
        )

        rows = quote.build_quote_rows(brief, [line])

        self.assertTrue(any(row and row[-1] == "Included" for row in rows))

    def test_layout_keeps_quantity_column_readable_and_signatory_title_visible(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        quantity_col = next(
            col
            for col in sheet.find(f"{NS_MAIN}cols")
            if col.attrib.get("min") == "2" and col.attrib.get("max") == "2"
        )
        self.assertGreaterEqual(float(quantity_col.attrib["width"]), 14.0)
        self.assertEqual(cell_value(sheet, "B123"), "Director")

    def test_layout_totals_use_bordered_styles(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        for ref in ("D93", "E93", "F93"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(border.find(f"{NS_MAIN}top").attrib.get("style"), "thin")

        for ref in ("D94", "E94", "F94"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(border.find(f"{NS_MAIN}top").attrib.get("style"), "thin")
            self.assertEqual(border.find(f"{NS_MAIN}bottom").attrib.get("style"), "double")

    def test_company_detail_drawing_is_wide_and_top_aligned(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            drawing = ET.fromstring(zf.read("xl/drawings/drawing1.xml"))

        text_anchor = next(
            anchor
            for anchor in drawing.findall(f"{NS_DRAWING}twoCellAnchor")
            if anchor.find(f"{NS_DRAWING}sp") is not None
        )
        from_col = int(text_anchor.find(f"{NS_DRAWING}from/{NS_DRAWING}col").text)
        to_col = int(text_anchor.find(f"{NS_DRAWING}to/{NS_DRAWING}col").text)
        body_pr = text_anchor.find(f"{NS_DRAWING}sp/{NS_DRAWING}txBody/{NS_A}bodyPr")

        self.assertGreater(to_col, from_col)
        self.assertEqual(body_pr.attrib.get("anchor"), "t")

    def test_excel_pdf_export_script_repairs_and_saves_workbook_before_export(self):
        script = quote.powershell_export_script(Path("quotation.xlsx"), Path("quotation.pdf"))

        self.assertIn("CorruptLoad", script)
        self.assertIn("$workbook.SaveAs($repairedWorkbookPath, 51)", script)
        self.assertIn("Move-Item -LiteralPath $repairedWorkbookPath -Destination $xlsxPath -Force", script)


if __name__ == "__main__":
    unittest.main()
