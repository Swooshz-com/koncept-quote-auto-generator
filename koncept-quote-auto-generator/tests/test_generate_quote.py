import csv
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


def font_for_style(styles_root, style_id):
    cell_xfs = styles_root.find(f"{NS_MAIN}cellXfs")
    fonts = styles_root.find(f"{NS_MAIN}fonts")
    font_id = int(cell_xfs[style_id].attrib.get("fontId", "0"))
    return fonts[font_id]


def alignment_for_style(styles_root, style_id):
    cell_xfs = styles_root.find(f"{NS_MAIN}cellXfs")
    return cell_xfs[style_id].find(f"{NS_MAIN}alignment")


def num_fmt_for_style(styles_root, style_id):
    cell_xfs = styles_root.find(f"{NS_MAIN}cellXfs")
    return cell_xfs[style_id].attrib.get("numFmtId")


def worksheet_formulas(root):
    return [formula.text or "" for formula in root.iter(f"{NS_MAIN}f")]


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

    def test_company_detail_drawing_starts_below_logo(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            drawing = ET.fromstring(zf.read("xl/drawings/drawing1.xml"))

        anchors = drawing.findall(f"{NS_DRAWING}twoCellAnchor")
        text_anchor = next(anchor for anchor in anchors if anchor.find(f"{NS_DRAWING}sp") is not None)
        logo_anchor = next(anchor for anchor in anchors if anchor.find(f"{NS_DRAWING}pic") is not None)
        text_from_row = int(text_anchor.find(f"{NS_DRAWING}from/{NS_DRAWING}row").text)
        logo_to_row = int(logo_anchor.find(f"{NS_DRAWING}to/{NS_DRAWING}row").text)
        text_ext = text_anchor.find(f"{NS_DRAWING}sp/{NS_DRAWING}spPr/{NS_A}xfrm/{NS_A}ext")
        text_off = text_anchor.find(f"{NS_DRAWING}sp/{NS_DRAWING}spPr/{NS_A}xfrm/{NS_A}off")
        logo_ext = logo_anchor.find(f"{NS_DRAWING}pic/{NS_DRAWING}spPr/{NS_A}xfrm/{NS_A}ext")
        logo_off = logo_anchor.find(f"{NS_DRAWING}pic/{NS_DRAWING}spPr/{NS_A}xfrm/{NS_A}off")
        text_left = int(text_off.attrib["x"])
        text_width = int(text_ext.attrib["cx"])
        logo_left = int(logo_off.attrib["x"])
        logo_width = int(logo_ext.attrib["cx"])

        self.assertGreater(text_from_row, logo_to_row)
        self.assertLessEqual(text_left, logo_left)
        self.assertGreaterEqual(text_left + text_width, logo_left + logo_width)
        self.assertLessEqual(text_width, 2600000)
        self.assertLessEqual(text_left + text_width, 7800000)

        run_sizes = [
            int(run_props.attrib["sz"])
            for run_props in text_anchor.findall(f".//{NS_A}rPr")
            if "sz" in run_props.attrib
        ]
        self.assertTrue(run_sizes)
        self.assertGreaterEqual(min(run_sizes), 550)

    def test_table_headers_bold_and_quantity_centered(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        for ref in ("A20", "B20", "C20"):
            self.assertIsNotNone(font_for_style(styles, cell_style(sheet, ref)).find(f"{NS_MAIN}b"))
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "B20")).attrib.get("horizontal"), "center")
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "B24")).attrib.get("horizontal"), "center")

    def test_price_cells_use_thousands_separator_number_format(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        self.assertEqual(num_fmt_for_style(styles, cell_style(sheet, "E24")), "4")
        self.assertEqual(num_fmt_for_style(styles, cell_style(sheet, "E93")), "4")
        self.assertEqual(num_fmt_for_style(styles, cell_style(sheet, "E94")), "4")

    def test_notes_are_plain_numbered_and_acceptance_text_stays_out_of_notes(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        self.assertEqual(cell_value(sheet, "A114"), "11.00")
        self.assertEqual(cell_value(sheet, "A115"), "12.00")
        self.assertIsNone(font_for_style(styles, cell_style(sheet, "A103")).find(f"{NS_MAIN}i"))
        self.assertIsNone(font_for_style(styles, cell_style(sheet, "A114")).find(f"{NS_MAIN}i"))
        self.assertEqual(cell_value(sheet, "E106"), "")
        self.assertEqual(cell_value(sheet, "E117"), "We accept the quotation amount and the terms")

    def test_excel_pdf_export_script_repairs_and_saves_workbook_before_export(self):
        script = quote.powershell_export_script(Path("quotation.xlsx"), Path("quotation.pdf"))

        self.assertIn("CorruptLoad", script)
        self.assertIn("$workbook.SaveAs($repairedWorkbookPath, 51)", script)
        self.assertIn("Move-Item -LiteralPath $repairedWorkbookPath -Destination $xlsxPath -Force", script)

    def test_brief_text_is_written_as_literal_xlsx_text_not_formulas(self):
        brief = {
            "company_identity": "Koncept Image",
            "quote_date": "2026-06-04",
            "client": {
                "name": "=WEBSERVICE(\"https://example.test/client\")",
                "attention": "Alex Tan",
            },
            "project": {
                "title": "=HYPERLINK(\"https://example.test/project\",\"project\")",
            },
            "line_items": [],
            "payment_terms": [
                "=WEBSERVICE(\"https://example.test/terms\")",
            ],
        }
        line = quote.QuoteLine(
            section="=WEBSERVICE(\"https://example.test/section\")",
            quantity=1,
            unit="lot",
            description="=HYPERLINK(\"https://example.test/line\",\"line\")",
            pricing_keyword="",
            display_price="=HYPERLINK(\"https://example.test/price\",\"price\")",
            matched_price=None,
            amount=None,
            match_status="manual-display",
            match_candidates=[],
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "quotation.xlsx"
            quote.write_quote_layout_xlsx(ROOT / "references" / "quotation-layout.xlsx", path, brief, [line])

            with zipfile.ZipFile(path) as zf:
                sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        self.assertEqual(worksheet_formulas(sheet), ["SUM(E22:E93)"])
        self.assertEqual(cell_value(sheet, "A6"), brief["client"]["name"])
        self.assertEqual(cell_value(sheet, "C22"), line.section)
        self.assertEqual(cell_value(sheet, "E24"), line.display_price)
        self.assertEqual(cell_value(sheet, "B100"), brief["payment_terms"][0])

    def test_csv_match_report_neutralizes_formula_leading_text(self):
        line = quote.QuoteLine(
            section="=WEBSERVICE(\"https://example.test/section\")",
            quantity=1,
            unit="lot",
            description="+HYPERLINK(\"https://example.test/line\",\"line\")",
            pricing_keyword="-10+20",
            display_price="",
            matched_price=None,
            amount=None,
            match_status="@manual",
            match_candidates=[],
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "pricing_matches.csv"
            quote.write_match_csv(path, [line])
            with path.open(newline="", encoding="utf-8") as f:
                rows = list(csv.reader(f))

        data = rows[1]
        self.assertEqual(data[0], "'@manual")
        self.assertEqual(data[1], "'=WEBSERVICE(\"https://example.test/section\")")
        self.assertEqual(data[2], "'+HYPERLINK(\"https://example.test/line\",\"line\")")
        self.assertEqual(data[3], "'-10+20")


if __name__ == "__main__":
    unittest.main()
