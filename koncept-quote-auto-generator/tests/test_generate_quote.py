import csv
import json
import re
import sys
import tempfile
import unittest
from unittest import mock
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import generate_quote as quote


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
NS_DC = "{http://purl.org/dc/elements/1.1/}"
NS_MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"


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


def cell(root, ref):
    for item in root.iter(f"{NS_MAIN}c"):
        if item.attrib.get("r") == ref:
            return item
    raise AssertionError(f"Cell {ref} was not written")


def cell_inline_runs(root, ref):
    inline = cell(root, ref).find(f"{NS_MAIN}is")
    if inline is None:
        return []
    runs = inline.findall(f"{NS_MAIN}r")
    if not runs:
        text = inline.find(f"{NS_MAIN}t")
        return [((text.text if text is not None else ""), False)]
    result = []
    for run in runs:
        text = "".join(text_node.text or "" for text_node in run.findall(f"{NS_MAIN}t"))
        run_props = run.find(f"{NS_MAIN}rPr")
        result.append((text, run_props is not None and run_props.find(f"{NS_MAIN}b") is not None))
    return result


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


def declared_xml_prefixes(xml_text, root_name):
    match = re.search(rf"<{root_name}\b[^>]*>", xml_text)
    if not match:
        raise AssertionError(f"Could not find <{root_name}> start tag")
    return set(re.findall(r"\sxmlns:([A-Za-z0-9_]+)=", match.group(0)))


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
    def test_pdf_generation_is_opt_in_by_default(self):
        with mock.patch.object(sys, "argv", ["generate_quote.py", "--brief", "brief.json"]):
            args = quote.parse_args()

        self.assertEqual(args.pdf_mode, "none")

    def test_markdown_pricing_source_is_default_template(self):
        with mock.patch.object(sys, "argv", ["generate_quote.py", "--brief", "brief.json"]):
            args = quote.parse_args()

        self.assertEqual(args.template, ROOT / "references" / "quotation-cost-template.md")

    def test_extract_price_rows_reads_authoritative_rag_markdown(self):
        markdown_path = ROOT / "references" / "quotation-cost-template.md"
        markdown_text = markdown_path.read_text(encoding="utf-8-sig")
        markdown_rows = quote.extract_price_rows(markdown_path)

        item_lines = [line for line in markdown_text.splitlines() if line.startswith("- **Item:**")]
        section_lines = [line for line in markdown_text.splitlines() if line.startswith("### ")]
        self.assertIn("kind: rag_pricing_source", markdown_text)
        self.assertIn("Clean RAG Pricing Source", markdown_text)
        self.assertIn("Customer-facing quotation line amount", markdown_text)
        self.assertIn("GST is shown separately in the quotation totals block", markdown_text)
        self.assertIn("### Floor Design", markdown_text)
        self.assertNotIn("### . Unsectioned", markdown_text)
        self.assertNotIn("source_workbook", markdown_text)
        self.assertNotIn("_Quotation Cost Template", markdown_text)
        self.assertNotIn("Anchor:", markdown_text)
        self.assertNotIn("| Row |", markdown_text)
        self.assertNotIn("| Aux N |", markdown_text)
        self.assertNotIn("| Aux O |", markdown_text)
        self.assertNotIn("| Formulas |", markdown_text)
        self.assertNotIn("| Section No |", markdown_text)
        self.assertIn("- **Default quantity:** 1.", markdown_text)
        self.assertIn("- **Default quote amount:** SGD 817.5.", markdown_text)
        self.assertEqual(len(section_lines), 11)
        self.assertEqual(len(item_lines), 103)
        self.assertEqual(len(markdown_rows), 103)
        self.assertTrue(all(row.description and row.cost > 0 and row.markup > 0 for row in markdown_rows))
        self.assertEqual([row.row_number for row in markdown_rows], list(range(1, 104)))
        self.assertIn('nos. 42" LED TV Monitor (With Speaker \u2013 Full HD).', markdown_text)
        self.assertEqual(markdown_rows[0].row_number, 1)
        self.assertEqual(markdown_rows[0].section, "Floor Design")
        self.assertEqual(markdown_rows[0].description, "m2 needle punch carpet in colour")
        self.assertEqual(markdown_rows[0].unit_hint, "sqm")
        self.assertEqual(markdown_rows[0].cost, 7)
        self.assertEqual(markdown_rows[0].gst_multiplier, 1.09)
        self.assertEqual(markdown_rows[0].markup, 1.5)
        self.assertEqual(markdown_rows[0].sale_unit_price, 10.5)
        wall_row = next(row for row in markdown_rows if row.description.startswith("m length single side partition wall at height 2.4m"))
        self.assertEqual(wall_row.description, "m length single side partition wall at height 2.4m; wooden construct in painted finished as per design proposal")
        self.assertEqual(wall_row.remark, "Backwall or any partition; PAINTED")
        acrylic_row = next(row for row in markdown_rows if "Sytem Profile partition" in row.description)
        self.assertEqual(acrylic_row.remark, "Acrylic System Partition. Extra calculations: 150+15+20; Octanorm")
        pe_row = next(row for row in markdown_rows if row.description == "Professional Engineer Endorsement for hanging")
        self.assertEqual(pe_row.remark, "PE. Extra values: 780; 68. Extra calculations: 90+690; 17*4; Extra values: 1000; 14.705882352941176. Extra calculations: 1000/68")
        boom_row = next(row for row in markdown_rows if row.description.startswith("Lot. rental of Boom Lift"))
        self.assertEqual(boom_row.cost, 500)
        self.assertEqual(boom_row.gst_multiplier, 1.09)
        self.assertEqual(boom_row.sale_unit_price, 750)

    def test_generated_styles_declares_ignorable_prefixes_without_excel_repair(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            styles_xml = zf.read("xl/styles.xml").decode("utf-8")
            styles = ET.fromstring(styles_xml)

        declared = declared_xml_prefixes(styles_xml, "styleSheet")
        ignorable = styles.attrib.get(NS_MC_IGNORABLE, "").split()

        self.assertTrue(ignorable)
        self.assertEqual([prefix for prefix in ignorable if prefix not in declared], [])

    def test_layout_workbook_scrubs_customer_visible_template_metadata(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            core_xml = zf.read("docProps/core.xml").decode("utf-8")
            core = ET.fromstring(core_xml)
            workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")

        creator = core.find(f"{NS_DC}creator")
        last_modified_by = core.find(f"{NS_CP}lastModifiedBy")
        workbook = ET.fromstring(workbook_xml)
        declared = declared_xml_prefixes(workbook_xml, "workbook")
        ignorable = workbook.attrib.get(NS_MC_IGNORABLE, "").split()

        self.assertEqual(creator.text, "Koncept Quote Auto-Generator")
        self.assertEqual(last_modified_by.text, "Koncept Quote Auto-Generator")
        self.assertNotIn("absPath", workbook_xml)
        self.assertNotIn("/Users/", workbook_xml)
        self.assertNotIn("Dropbox", workbook_xml)
        self.assertEqual([prefix for prefix in ignorable if prefix not in declared], [])
        self.assertNotIn("<mc:Choice Requires=\"x15\" />", workbook_xml)

    def test_default_generation_removes_stale_pdf_output(self):
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
            "line_items": [
                {
                    "section": "Furniture",
                    "quantity": 1,
                    "unit": "lot",
                    "description": "Furniture package",
                    "display_price": "Included",
                }
            ],
        }
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            brief_path = tmp_path / "brief.json"
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            stale_pdf = out_dir / "quotation.pdf"
            stale_pdf.write_bytes(b"old pdf")
            brief_path.write_text(json.dumps(brief), encoding="utf-8")

            with mock.patch.object(sys, "argv", ["generate_quote.py", "--brief", str(brief_path), "--out", str(out_dir)]):
                self.assertEqual(quote.main(), 0)

            self.assertTrue((out_dir / "quotation.xlsx").exists())
            self.assertFalse(stale_pdf.exists())

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

        self.assertEqual(cell_value(sheet, "D92"), "Total")
        self.assertEqual(cell_value(sheet, "D93"), "GST 9%")
        self.assertEqual(cell_value(sheet, "D94"), "Total including GST")
        self.assertEqual(cell_value(sheet, "F92"), "SGD")
        self.assertEqual(cell_value(sheet, "F93"), "SGD")
        self.assertEqual(cell_value(sheet, "F94"), "SGD")
        self.assertEqual(worksheet_formulas(sheet), ["SUM(E22:E91)", "ROUND(E92*0.090000,0)", "SUM(E92:E93)"])

        for ref in ("D92", "E92", "F92"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(border.find(f"{NS_MAIN}top").attrib.get("style"), "thin")

        for ref in ("D93", "E93", "F93"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertIsNone(border.find(f"{NS_MAIN}top").attrib.get("style"))
            self.assertIsNone(border.find(f"{NS_MAIN}bottom").attrib.get("style"))

        for ref in ("D94", "E94", "F94"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(border.find(f"{NS_MAIN}top").attrib.get("style"), "thin")
            self.assertEqual(border.find(f"{NS_MAIN}bottom").attrib.get("style"), "double")

    def test_default_terms_match_sample_text_and_bold_emphasis(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        self.assertEqual(cell_value(sheet, "A99"), "Terms & Conditions :")
        self.assertIsNotNone(font_for_style(styles, cell_style(sheet, "A99")).find(f"{NS_MAIN}b"))
        self.assertEqual(cell_value(sheet, "A100"), "1.00")
        self.assertEqual(cell_value(sheet, "B100"), "70% payment upon confirmation and signing of contract.")
        self.assertEqual(cell_inline_runs(sheet, "B100"), [("70% payment upon confirmation and signing of contract.", True)])
        self.assertEqual(cell_value(sheet, "A101"), "2.00")
        self.assertEqual(cell_value(sheet, "B101"), "30% balance upon handover before show starts")
        self.assertEqual(cell_inline_runs(sheet, "B101"), [("30% balance upon handover before show starts", True)])
        self.assertEqual(cell_value(sheet, "B102"), "All cheques should be crossed and made payable to Koncept Image Pte Ltd")
        self.assertEqual(
            cell_inline_runs(sheet, "B102"),
            [
                ("All cheques should be crossed and made payable to ", False),
                ("Koncept Image Pte Ltd", True),
            ],
        )

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
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))

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
        text_top = int(text_off.attrib["y"])
        text_width = int(text_ext.attrib["cx"])
        logo_left = int(logo_off.attrib["x"])
        logo_top = int(logo_off.attrib["y"])
        logo_width = int(logo_ext.attrib["cx"])
        logo_height = int(logo_ext.attrib["cy"])
        text_from_col = int(text_anchor.find(f"{NS_DRAWING}from/{NS_DRAWING}col").text)
        text_to_col = int(text_anchor.find(f"{NS_DRAWING}to/{NS_DRAWING}col").text)
        text_from_col_off = int(text_anchor.find(f"{NS_DRAWING}from/{NS_DRAWING}colOff").text)
        logo_from_col_off = int(logo_anchor.find(f"{NS_DRAWING}from/{NS_DRAWING}colOff").text)
        logo_to_col_off = int(logo_anchor.find(f"{NS_DRAWING}to/{NS_DRAWING}colOff").text)
        logo_from_row_off = int(logo_anchor.find(f"{NS_DRAWING}from/{NS_DRAWING}rowOff").text)
        logo_to_row_off = int(logo_anchor.find(f"{NS_DRAWING}to/{NS_DRAWING}rowOff").text)

        print_titles = workbook.find(f"{NS_MAIN}definedNames/{NS_MAIN}definedName[@name='_xlnm.Print_Titles']")
        self.assertIsNotNone(print_titles)
        self.assertEqual(print_titles.text, "Quotation!$1:$3")
        self.assertGreater(text_from_row, logo_to_row)
        self.assertEqual(text_from_row, 3)
        self.assertLessEqual(text_top - (logo_top + logo_height), 220000)
        self.assertLessEqual(text_left, logo_left)
        self.assertGreaterEqual(text_left + text_width, logo_left + logo_width)
        self.assertGreaterEqual(logo_width, 2950000)
        self.assertGreaterEqual(logo_height, 630000)
        self.assertLessEqual(logo_from_col_off, 0)
        self.assertGreaterEqual(logo_to_col_off, 1720000)
        self.assertLessEqual(logo_from_row_off, 0)
        self.assertGreaterEqual(logo_to_row_off, 410000)
        self.assertGreaterEqual(text_width, 3300000)
        self.assertLessEqual(text_width, 3500000)
        self.assertEqual(text_from_col, 7)
        self.assertGreaterEqual(text_to_col, 9)
        self.assertEqual(text_from_col_off, 0)
        self.assertEqual(text_left, logo_left)

        paragraph_alignments = [
            paragraph.find(f"{NS_A}pPr").attrib.get("algn")
            for paragraph in text_anchor.findall(f"{NS_DRAWING}sp/{NS_DRAWING}txBody/{NS_A}p")
            if paragraph.find(f"{NS_A}pPr") is not None
        ]
        self.assertTrue(paragraph_alignments)
        self.assertEqual(set(paragraph_alignments), {"l"})

        paragraphs = [
            "".join(text_node.text or "" for text_node in paragraph.findall(f".//{NS_A}t"))
            for paragraph in text_anchor.findall(f"{NS_DRAWING}sp/{NS_DRAWING}txBody/{NS_A}p")
        ]
        self.assertEqual(
            paragraphs,
            [
                "Koncept Image Pte Limited",
                "61 Kaki Bukit Ave 1, #02-26,",
                "Shunli Industrial Park",
                "Singapore 417943",
                "Telephone: +6568177477",
                "",
                "Bank Detail:",
                "United Overseas Bank Limited, 80",
                "Raffles Place",
                "Singapore 048624",
                "Account: 335-3020-445",
                "Swift Code: UOVBSGSG",
                "",
                "Project No: KI-TEST-001",
            ],
        )
        bank_index = paragraphs.index("Bank Detail:")
        self.assertEqual(paragraphs[bank_index - 1], "")
        project_index = paragraphs.index("Project No: KI-TEST-001")
        self.assertEqual(paragraphs[project_index - 1], "")
        self.assertFalse(any(current == "" and following == "" for current, following in zip(paragraphs, paragraphs[1:])))

        run_sizes = [
            int(run_props.attrib["sz"])
            for run_props in text_anchor.findall(f".//{NS_A}rPr")
            if "sz" in run_props.attrib
        ]
        self.assertTrue(run_sizes)
        self.assertEqual(min(run_sizes), 900)
        self.assertEqual(max(run_sizes), 900)

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

        self.assertEqual(worksheet_formulas(sheet), ["SUM(E22:E91)", "SUM(E92:E93)"])
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
