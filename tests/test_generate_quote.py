import base64
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
KONCEPT_PROFILE = ROOT / "profiles" / "koncept"
KONCEPT_CATALOG = ROOT / "pricing-references" / "koncept" / "pricing-catalog.json"
KONCEPT_LAYOUT = KONCEPT_PROFILE / "quotation-layout.xlsx"
KONCEPT_LOGO = KONCEPT_PROFILE / "assets" / "koncept-header-logo.jpeg"
sys.path.insert(0, str(ROOT / "scripts"))

import generate_quote as quote
import build_pricing_catalog as pricing_catalog


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
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


def find_cell_ref(root, expected):
    for cell in root.iter(f"{NS_MAIN}c"):
        if cell_value(root, cell.attrib.get("r", "")) == expected:
            return cell.attrib.get("r", "")
    return ""


def find_cell_refs(root, expected):
    return [
        cell.attrib.get("r", "")
        for cell in root.iter(f"{NS_MAIN}c")
        if cell_value(root, cell.attrib.get("r", "")) == expected
    ]


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
        return [((text.text if text is not None else ""), False, False, False)]
    result = []
    for run in runs:
        text = "".join(text_node.text or "" for text_node in run.findall(f"{NS_MAIN}t"))
        run_props = run.find(f"{NS_MAIN}rPr")
        bold = run_props.find(f"{NS_MAIN}b") if run_props is not None else None
        result.append(
            (
                text,
                bold is not None and bold.attrib.get("val", "1") not in {"0", "false", "False"},
                run_props is not None and run_props.find(f"{NS_MAIN}i") is not None,
                run_props is not None and run_props.find(f"{NS_MAIN}u") is not None,
            )
        )
    return result


def drawing_paragraph_runs(root):
    paragraphs = []
    for paragraph in root.findall(f".//{NS_A}p"):
        runs = []
        for run in paragraph.findall(f"{NS_A}r"):
            text = "".join(text_node.text or "" for text_node in run.findall(f"{NS_A}t"))
            run_props = run.find(f"{NS_A}rPr")
            runs.append(
                (
                    text,
                    run_props is not None and run_props.attrib.get("b") == "1",
                    run_props is not None and run_props.attrib.get("i") == "1",
                    run_props is not None and run_props.attrib.get("u") == "sng",
                )
            )
        if runs:
            paragraphs.append(runs)
    return paragraphs


def drawing_with_picture_xml(name="Product Screenshot", rel_id="rId5"):
    root = ET.Element(f"{NS_DRAWING}wsDr")
    anchor = ET.SubElement(root, f"{NS_DRAWING}twoCellAnchor")
    from_marker = ET.SubElement(anchor, f"{NS_DRAWING}from")
    for tag, value in (("col", "0"), ("colOff", "0"), ("row", "0"), ("rowOff", "0")):
        ET.SubElement(from_marker, f"{NS_DRAWING}{tag}").text = value
    to_marker = ET.SubElement(anchor, f"{NS_DRAWING}to")
    for tag, value in (("col", "1"), ("colOff", "0"), ("row", "1"), ("rowOff", "0")):
        ET.SubElement(to_marker, f"{NS_DRAWING}{tag}").text = value
    pic = ET.SubElement(anchor, f"{NS_DRAWING}pic")
    nv_pic_pr = ET.SubElement(pic, f"{NS_DRAWING}nvPicPr")
    ET.SubElement(nv_pic_pr, f"{NS_DRAWING}cNvPr", {"id": "9", "name": name})
    ET.SubElement(nv_pic_pr, f"{NS_DRAWING}cNvPicPr")
    blip_fill = ET.SubElement(pic, f"{NS_DRAWING}blipFill")
    ET.SubElement(blip_fill, f"{NS_A}blip", {f"{NS_REL}embed": rel_id})
    ET.SubElement(pic, f"{NS_DRAWING}spPr")
    ET.SubElement(anchor, f"{NS_DRAWING}clientData")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def empty_drawing_xml():
    return ET.tostring(ET.Element(f"{NS_DRAWING}wsDr"), encoding="utf-8", xml_declaration=True)


def drawing_rels_xml(*relationships):
    root = ET.Element(f"{NS_PACKAGE_REL}Relationships")
    for rel_id, target in relationships:
        ET.SubElement(
            root,
            f"{NS_PACKAGE_REL}Relationship",
            {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": target,
            },
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def empty_content_types_xml():
    return (
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" />'
    )


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


def num_fmt_code_for_style(styles_root, style_id):
    num_fmt_id = num_fmt_for_style(styles_root, style_id)
    num_fmts = styles_root.find(f"{NS_MAIN}numFmts")
    if num_fmts is not None:
        for num_fmt in num_fmts.findall(f"{NS_MAIN}numFmt"):
            if num_fmt.attrib.get("numFmtId") == num_fmt_id:
                return num_fmt.attrib.get("formatCode")
    return num_fmt_id


def font_color(font):
    color = font.find(f"{NS_MAIN}color")
    return color.attrib if color is not None else {}


def worksheet_formulas(root):
    return [formula.text or "" for formula in root.iter(f"{NS_MAIN}f")]


def defined_name_text(workbook, name):
    defined_name = workbook.find(f"{NS_MAIN}definedNames/{NS_MAIN}definedName[@name='{name}']")
    return defined_name.text if defined_name is not None else ""


def row_break_ids(sheet):
    row_breaks = sheet.find(f"{NS_MAIN}rowBreaks")
    if row_breaks is None:
        return []
    return [int(brk.attrib["id"]) for brk in row_breaks.findall(f"{NS_MAIN}brk")]


def manual_print_page_for_row(row_number):
    if row_number <= quote.FIRST_PRINT_PAGE_END_ROW:
        return 1
    return 2 + ((row_number - quote.CONTINUATION_PAGE_START_ROW) // quote.CONTINUATION_PAGE_HEIGHT)


def declared_xml_prefixes(xml_text, root_name):
    match = re.search(rf"<{root_name}\b[^>]*>", xml_text)
    if not match:
        raise AssertionError(f"Could not find <{root_name}> start tag")
    return set(re.findall(r"\sxmlns:([A-Za-z0-9_]+)=", match.group(0)))


def logo_data_url():
    return f"data:image/jpeg;base64,{base64.b64encode(KONCEPT_LOGO.read_bytes()).decode('ascii')}"


def generate_layout_workbook(brief_updates=None):
    brief = {
        "company_identity": "Koncept Image",
        "quote_date": "2026-06-04",
        "project_number": "KI-TEST-001",
        "client": {
            "name": "Sample Client",
            "attention": "Alex Tan",
            "address": ["10 Sample Street", "#02-03 Sample Building"],
        },
        "project": {
            "title": "Sample Project",
        },
        "line_items": [],
        "payment_terms": [],
        "company": {
            "name": "Koncept Image Pte Ltd",
            "header_lines": ["Koncept Image Pte Limited", "61 Kaki Bukit Ave 1"],
            "logo_data_url": logo_data_url(),
        },
        "acceptance": {
            "company_name": "Koncept Image Pte Ltd",
            "text": "We accept the quotation amount and the terms",
            "person_label": "Person in charge",
            "stamp_label": "Company name & stamp",
            "date_label": "Date:",
        },
        "signature": {
            "koncept_signatory": "Francies Cheng",
            "koncept_title": "Director",
            "koncept_date_label": "Date:",
        },
    }
    if brief_updates:
        brief.update(brief_updates)
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
    quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, [line])
    return tmp, path


class GenerateQuoteRowsTest(unittest.TestCase):
    def test_pdf_generation_is_opt_in_by_default(self):
        with mock.patch.object(sys, "argv", ["generate_quote.py", "--brief", "brief.json"]):
            args = quote.parse_args()

        self.assertEqual(args.pdf_mode, "none")

    def test_json_pricing_catalog_is_default_template(self):
        with mock.patch.object(sys, "argv", ["generate_quote.py", "--brief", "brief.json"]):
            args = quote.parse_args()

        self.assertEqual(args.template, KONCEPT_CATALOG)

    def test_xlsx_catalog_builder_collapses_continuation_rows(self):
        rows = [
            [1, None, "Booth Structure"],
            [None, None, "m length single side partition wall at height 2.4m", None, None, 0, None, 180, None, 1.5, None, "Backwall or any partition"],
            [None, None, "wooden construct in painted finished as per design proposal", None, None, None, None, None, None, None, None, "PAINTED"],
            [None, None, "sets of 3D backlit lettering", None, None, 0, None, 400, 1.09, 1.5, None, "Backlit Lettering"],
            [None, None, None, None, None, None, None, None, None, None, None, "NOTE: Per set max 2m long"],
        ]
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            quote.write_minimal_xlsx(source, rows)

            catalog = pricing_catalog.build_catalog(source, source_label="source.xlsx")

        self.assertEqual(catalog["schema_version"], 1)
        self.assertNotIn("source", catalog)
        self.assertEqual(len(catalog["items"]), 2)

        wall = catalog["items"][0]
        self.assertEqual(wall["id"], "booth-structure.single-side-partition-wall-at-height-2-4m-wooden-construct-in-painted-finished-as-per-design-proposal")
        self.assertNotIn("source_row", wall)
        self.assertEqual(wall["section"], "Booth Structure")
        self.assertEqual(wall["description"], "m length single side partition wall at height 2.4m; wooden construct in painted finished as per design proposal")
        self.assertEqual(wall["remarks"], ["Backwall or any partition", "PAINTED"])
        self.assertEqual(wall["unit_hint"], "m length")
        self.assertIn("single side partition wall at height 2.4m", wall["aliases"])

        lettering = catalog["items"][1]
        self.assertEqual(lettering["id"], "booth-structure.3d-backlit-lettering")
        self.assertEqual(lettering["remarks"], ["Backlit Lettering", "NOTE: Per set max 2m long"])

        ai_reference_markdown = pricing_catalog.catalog_to_ai_reference_markdown(catalog)
        self.assertNotIn("source row", ai_reference_markdown)
        self.assertIn("Pricing Catalog AI Reference", ai_reference_markdown)
        self.assertIn("m length single side partition wall at height 2.4m; wooden construct in painted finished as per design proposal", ai_reference_markdown)
        self.assertIn("Backwall or any partition; PAINTED", ai_reference_markdown)

    def test_extract_price_rows_reads_json_catalog(self):
        catalog_path = KONCEPT_CATALOG
        rows = quote.extract_price_rows(catalog_path)

        self.assertGreaterEqual(len(rows), 100)
        self.assertEqual(rows[0].pricing_id, "floor-design.needle-punch-carpet-in-colour")
        self.assertEqual(rows[0].section, "Floor Design")
        self.assertEqual(rows[0].description, "m2 needle punch carpet in colour")
        self.assertEqual(rows[0].unit_hint, "sqm")
        self.assertEqual(rows[0].cost, 7)
        self.assertEqual(rows[0].gst_multiplier, 1.09)
        self.assertEqual(rows[0].markup, 1.5)
        self.assertIn("needle punch", rows[0].aliases)

        wall_row = next(row for row in rows if row.pricing_id == "booth-structure.single-side-partition-wall-at-height-2-4m-wooden-construct-in-painted-finished-as-per-design-proposal")
        self.assertEqual(wall_row.description, "m length single side partition wall at height 2.4m; wooden construct in painted finished as per design proposal")
        self.assertEqual(wall_row.remark, "Backwall or any partition; PAINTED")

        graphics_row = next(row for row in rows if row.pricing_id == "graphics.vinyl-printed-graphics")
        self.assertEqual(graphics_row.description, "m2 of vinyl printed graphics")
        self.assertIn("printed graphics on wall", graphics_row.remark.lower())

    def test_catalog_id_pricing_keyword_matches_exact_catalog_item(self):
        rows = quote.extract_price_rows(KONCEPT_CATALOG)

        status, match, _ = quote.find_price_match("graphics.vinyl-printed-graphics", rows)

        self.assertEqual(status, "matched")
        self.assertIsNotNone(match)
        self.assertEqual(match.pricing_id, "graphics.vinyl-printed-graphics")

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

        self.assertEqual(creator.text, "Swooshz Quote Generator")
        self.assertEqual(last_modified_by.text, "Swooshz Quote Generator")
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

    def test_unresolved_manual_display_placeholder_blocks_quotation_output(self):
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
                    "section": "Furniture Rental",
                    "quantity": 4,
                    "unit": "nos",
                    "description": "Round cafe tables for seating clusters",
                    "display_price": "Manual display price",
                }
            ],
        }
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            brief_path = tmp_path / "brief.json"
            out_dir = tmp_path / "out"
            brief_path.write_text(json.dumps(brief), encoding="utf-8")

            with mock.patch.object(sys, "argv", ["generate_quote.py", "--brief", str(brief_path), "--out", str(out_dir)]):
                self.assertEqual(quote.main(), 2)

            self.assertTrue((out_dir / "pricing_matches.csv").exists())
            self.assertFalse((out_dir / "quotation.xlsx").exists())

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

    def test_unresolved_manual_display_rows_require_confirmation(self):
        unresolved = quote.QuoteLine(
            section="Furniture Rental",
            quantity=4,
            unit="nos",
            description="Round cafe tables for seating clusters",
            pricing_keyword="",
            display_price="",
            matched_price=None,
            amount=None,
            match_status="manual-display",
            match_candidates=[],
        )
        placeholder = quote.QuoteLine(
            section="Furniture Rental",
            quantity=4,
            unit="nos",
            description="Round cafe tables for seating clusters",
            pricing_keyword="",
            display_price="Manual display price",
            matched_price=None,
            amount=None,
            match_status="manual-display",
            match_candidates=[],
        )
        included = quote.QuoteLine(
            section="Furniture Rental",
            quantity=1,
            unit="lot",
            description="Furniture package included in package",
            pricing_keyword="",
            display_price="Included",
            matched_price=None,
            amount=None,
            match_status="manual-display",
            match_candidates=[],
        )

        issues = quote.confirmation_issues([], [unresolved])
        placeholder_issues = quote.confirmation_issues([], [placeholder])
        confirmed_issues = quote.confirmation_issues([], [included])

        self.assertIn(
            "Manual display pricing required: Round cafe tables for seating clusters / enter a display price, choose a catalog keyword, or remove this line",
            issues,
        )
        self.assertIn(
            "Manual display pricing required: Round cafe tables for seating clusters / enter a display price, choose a catalog keyword, or remove this line",
            placeholder_issues,
        )
        self.assertEqual(confirmed_issues, [])

    def test_structure_sections_are_itemized_by_default(self):
        lines = [
            quote.QuoteLine(
                section="Booth Structure",
                quantity=2,
                unit="m length",
                description="Painted partition wall",
                pricing_keyword="booth-structure.partition-wall",
                display_price="",
                matched_price=None,
                amount=300,
                match_status="matched",
                match_candidates=[],
            ),
            quote.QuoteLine(
                section="Booth Structure",
                quantity=1,
                unit="each",
                description="Painted support column",
                pricing_keyword="booth-structure.support-column",
                display_price="",
                matched_price=None,
                amount=120,
                match_status="matched",
                match_candidates=[],
            ),
        ]

        entries = quote.render_quote_entries(lines, {})

        self.assertEqual(entries[0]["kind"], "section")
        self.assertIsNone(entries[0].get("amount"))
        self.assertNotIn("coverage", entries[0])
        self.assertEqual(entries[1]["amount"], 300)
        self.assertEqual(entries[2]["amount"], 120)

    def test_explicit_section_total_mode_still_groups_when_requested(self):
        lines = [
            quote.QuoteLine(
                section="Booth Structure",
                quantity=2,
                unit="m length",
                description="Painted partition wall",
                pricing_keyword="booth-structure.partition-wall",
                display_price="",
                matched_price=None,
                amount=300,
                match_status="matched",
                match_candidates=[],
            ),
            quote.QuoteLine(
                section="Booth Structure",
                quantity=1,
                unit="each",
                description="Painted support column",
                pricing_keyword="booth-structure.support-column",
                display_price="",
                matched_price=None,
                amount=120,
                match_status="matched",
                match_candidates=[],
            ),
        ]

        entries = quote.render_quote_entries(lines, {"section_pricing": {"Booth Structure": "section_total"}})

        self.assertEqual(entries[0]["amount"], 420)
        self.assertEqual(entries[1]["amount"], None)
        self.assertEqual(entries[2]["amount"], None)
        self.assertEqual(entries[0]["coverage"], "Covers line items 1.1 to 1.2")

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
        self.assertTrue(find_cell_ref(sheet, "Director").startswith("B"))

    def test_layout_totals_use_bordered_styles(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        total_ref = find_cell_ref(sheet, "Total")
        gst_ref = find_cell_ref(sheet, "GST 9%")
        grand_ref = find_cell_ref(sheet, "Total including GST")
        self.assertEqual(total_ref, "D27")
        self.assertEqual(gst_ref, "D28")
        self.assertEqual(grand_ref, "D29")
        self.assertEqual(cell_value(sheet, "F27"), "SGD")
        self.assertEqual(cell_value(sheet, "F28"), "SGD")
        self.assertEqual(cell_value(sheet, "F29"), "SGD")
        self.assertEqual(cell_value(sheet, "D92"), "")
        self.assertEqual(worksheet_formulas(sheet), ["SUM(E22:E26)", "ROUND(E27*0.090000,0)", "SUM(E27:E28)"])
        self.assertAlmostEqual(float(cell_value(sheet, "E27")), 2400.0)
        self.assertAlmostEqual(float(cell_value(sheet, "E28")), 216.0)
        self.assertAlmostEqual(float(cell_value(sheet, "E29")), 2616.0)

        for ref in ("D27", "E27", "F27"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(border.find(f"{NS_MAIN}top").attrib.get("style"), "thin")

        for ref in ("D28", "E28", "F28"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertIsNone(border.find(f"{NS_MAIN}top").attrib.get("style"))
            self.assertIsNone(border.find(f"{NS_MAIN}bottom").attrib.get("style"))

        for ref in ("D29", "E29", "F29"):
            border = border_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(border.find(f"{NS_MAIN}top").attrib.get("style"), "thin")
            self.assertEqual(border.find(f"{NS_MAIN}bottom").attrib.get("style"), "double")

        cols = sheet.find(f"{NS_MAIN}cols")
        d_col = next(col for col in cols if col.attrib.get("min") == "4" and col.attrib.get("max") == "4")
        self.assertGreaterEqual(float(d_col.attrib["width"]), 20.0)
        calc_pr = workbook.find(f"{NS_MAIN}calcPr")
        self.assertIsNotNone(calc_pr)
        self.assertEqual(calc_pr.attrib.get("fullCalcOnLoad"), "1")
        self.assertEqual(calc_pr.attrib.get("forceFullCalc"), "1")

    def test_layout_uses_dynamic_print_area_without_forced_extra_pages(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            workbook = ET.fromstring(zf.read("xl/workbook.xml"))

        dimension = sheet.find(f"{NS_MAIN}dimension")
        row_breaks = sheet.find(f"{NS_MAIN}rowBreaks")
        sheet_view = sheet.find(f"{NS_MAIN}sheetViews/{NS_MAIN}sheetView")

        self.assertEqual(defined_name_text(workbook, "_xlnm.Print_Area"), "Quotation!$A$1:$I$40")
        self.assertEqual(dimension.attrib.get("ref"), "A1:I40")
        self.assertIsNone(row_breaks)
        self.assertNotEqual(sheet_view.attrib.get("view"), "pageBreakPreview")
        self.assertEqual(cell_value(sheet, "B38"), "Director")
        self.assertEqual(cell_value(sheet, "B40"), "")
        self.assertFalse([
            item.attrib.get("r")
            for item in sheet.iter(f"{NS_MAIN}c")
            if quote.parse_cell_ref(item.attrib.get("r", "A1"))[0] > 40 and quote.parse_cell_ref(item.attrib.get("r", "A1"))[1] <= 9
        ])

    def test_layout_extends_quote_table_past_preserved_second_page(self):
        brief = {
            "company_identity": "Koncept Image",
            "quote_date": "2026-06-04",
            "client": {
                "name": "Sample Client",
                "attention": "Alex Tan",
            },
            "project": {
                "title": "RE: Large Generated Booth",
            },
            "line_items": [],
            "payment_terms": [],
            "signature": {
                "koncept_signatory": "Francies Cheng",
                "koncept_title": "Director",
            },
        }
        price = quote.PriceRow(1, "Generated", "generated booth component", "lot", 100, 1.09, 1, "")
        lines = [
            quote.QuoteLine(
                section=f"Generated Section {(index - 1) // 4 + 1}",
                quantity=1,
                unit="lot",
                description=f"Generated booth component {index}",
                pricing_keyword="generated booth component",
                display_price="",
                matched_price=price,
                amount=100,
                match_status="matched",
                match_candidates=[],
            )
            for index in range(1, 41)
        ]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "quotation.xlsx"
            quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, lines)

            with zipfile.ZipFile(path) as zf:
                sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
                workbook = ET.fromstring(zf.read("xl/workbook.xml"))

        last_item_ref = find_cell_ref(sheet, "Generated booth component 40")
        total_ref = find_cell_ref(sheet, "Total")
        last_item_row = quote.parse_cell_ref(last_item_ref)[0]
        total_row = quote.parse_cell_ref(total_ref)[0]

        self.assertGreater(last_item_row, 91)
        self.assertGreater(total_row, last_item_row)
        self.assertEqual(cell_value(sheet, f"F{total_row}"), "SGD")
        self.assertEqual(defined_name_text(workbook, "_xlnm.Print_Area"), f"Quotation!$A$1:$I${total_row + 13}")
        self.assertEqual(row_break_ids(sheet)[:2], [61, 122])

    def test_layout_uses_manual_continuation_pages_with_table_headers_only(self):
        brief = {
            "company_identity": "Koncept Image",
            "quote_date": "2026-06-04",
            "client": {
                "name": "Sample Client",
                "attention": "Alex Tan",
            },
            "project": {
                "title": "RE: Large Generated Booth",
            },
            "line_items": [],
            "payment_terms": [],
            "signature": {
                "koncept_signatory": "Francies Cheng",
                "koncept_title": "Director",
            },
        }
        price = quote.PriceRow(1, "Generated", "generated booth component", "lot", 100, 1.09, 1, "")
        lines = [
            quote.QuoteLine(
                section=f"Generated Section {(index - 1) // 4 + 1}",
                quantity=1,
                unit="lot",
                description=f"Generated booth component {index}",
                pricing_keyword="generated booth component",
                display_price="",
                matched_price=price,
                amount=100,
                match_status="matched",
                match_candidates=[],
            )
            for index in range(1, 41)
        ]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "quotation.xlsx"
            quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, lines)

            with zipfile.ZipFile(path) as zf:
                sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
                styles = ET.fromstring(zf.read("xl/styles.xml"))

        header_refs = find_cell_refs(sheet, "Pos.")
        title_refs = find_cell_refs(sheet, "RE: Large Generated Booth")

        self.assertEqual(header_refs[:3], ["A20", "A64", "A125"])
        self.assertEqual(title_refs, ["A18"])
        self.assertEqual(row_break_ids(sheet)[:2], [61, 122])
        for ref in ("E20", "E21", "E64", "E65"):
            self.assertEqual(alignment_for_style(styles, cell_style(sheet, ref)).attrib.get("horizontal"), "right")

    def test_layout_keeps_totals_together_on_one_manual_page(self):
        brief = {
            "company_identity": "Koncept Image",
            "quote_date": "2026-06-04",
            "client": {
                "name": "Sample Client",
                "attention": "Alex Tan",
            },
            "project": {
                "title": "Totals Boundary Booth",
            },
            "line_items": [],
            "payment_terms": [],
            "signature": {
                "koncept_signatory": "Francies Cheng",
                "koncept_title": "Director",
            },
        }
        price = quote.PriceRow(1, "Generated", "generated booth component", "lot", 100, 1.09, 1, "")
        lines = [
            quote.QuoteLine(
                section=f"Generated Section {(index - 1) // 4 + 1}",
                quantity=1,
                unit="lot",
                description=f"Generated booth component {index}",
                pricing_keyword="generated booth component",
                display_price="",
                matched_price=price,
                amount=100,
                match_status="matched",
                match_candidates=[],
            )
            for index in range(1, 41)
        ]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "quotation.xlsx"
            quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, lines)

            with zipfile.ZipFile(path) as zf:
                sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        total_row = quote.parse_cell_ref(find_cell_ref(sheet, "Total"))[0]
        gst_row = quote.parse_cell_ref(find_cell_ref(sheet, "GST 9%"))[0]
        grand_row = quote.parse_cell_ref(find_cell_ref(sheet, "Total including GST"))[0]

        self.assertEqual(
            {manual_print_page_for_row(total_row), manual_print_page_for_row(gst_row), manual_print_page_for_row(grand_row)},
            {manual_print_page_for_row(total_row)},
        )

    def test_empty_terms_and_notes_do_not_insert_default_rows(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        self.assertEqual(find_cell_ref(sheet, "Terms & Conditions:"), "")
        self.assertEqual(find_cell_ref(sheet, "Note:"), "")
        self.assertEqual(find_cell_ref(sheet, "All cheques should be crossed and made payable to Koncept Image Pte Ltd"), "")
        self.assertEqual(
            find_cell_ref(sheet, "The above contract does not include application fees to any relevant authorities and electrical connection fees."),
            "",
        )
        self.assertEqual(cell_value(sheet, "B32"), "Koncept Image Pte Ltd")
        self.assertEqual(cell_value(sheet, "E32"), "We accept the quotation amount and the terms")

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
            media_names = set(zf.namelist())

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
                "61 Kaki Bukit Ave 1",
                "",
                "Project No: KI-TEST-001",
            ],
        )
        project_index = paragraphs.index("Project No: KI-TEST-001")
        self.assertIn("xl/media/header_logo.jpeg", media_names)
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

        for ref in ("A20", "B20", "C20", "E20"):
            self.assertIsNotNone(font_for_style(styles, cell_style(sheet, ref)).find(f"{NS_MAIN}b"))
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "B20")).attrib.get("horizontal"), "center")
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "E20")).attrib.get("horizontal"), "right")
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "E21")).attrib.get("horizontal"), "right")
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "B24")).attrib.get("horizontal"), "center")

    def test_client_address_rows_are_left_aligned(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        self.assertEqual(cell_value(sheet, "A7"), "10 Sample Street")
        self.assertEqual(cell_value(sheet, "A8"), "#02-03 Sample Building")
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "A7")).attrib.get("horizontal"), "left")
        self.assertEqual(alignment_for_style(styles, cell_style(sheet, "A8")).attrib.get("horizontal"), "left")

    def test_attention_contact_name_and_title_are_split_from_label(self):
        tmp, path = generate_layout_workbook({
            "client": {
                "name": "Sample Client",
                "attention": "Melissa Ong",
                "title": "Senior Event Producer",
                "address": ["10 Sample Street", "#02-03 Sample Building"],
            },
        })
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        self.assertEqual(cell_value(sheet, "A11"), "Attention:")
        self.assertEqual(cell_value(sheet, "B12"), "Melissa Ong")
        self.assertEqual(cell_value(sheet, "B13"), "Senior Event Producer")
        self.assertEqual(cell_value(sheet, "A12"), "")
        self.assertEqual(cell_value(sheet, "A14"), "")
        self.assertEqual(cell_value(sheet, "A15"), "")
        self.assertEqual(cell_value(sheet, "A16"), str(quote.excel_date_serial("2026-06-04")))
        self.assertEqual(cell_inline_runs(sheet, "A11"), [("Attention:", True, False, False)])
        self.assertEqual(cell_inline_runs(sheet, "B12"), [("Melissa Ong", True, False, False)])
        self.assertEqual(cell_inline_runs(sheet, "B13"), [("Senior Event Producer", False, False, False)])

    def test_quote_date_cell_uses_customer_facing_date_format(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        self.assertEqual(cell_value(sheet, "A16"), str(quote.excel_date_serial("2026-06-04")))
        self.assertEqual(num_fmt_code_for_style(styles, cell_style(sheet, "A16")).replace("\\ ", " "), "dd mmmm yyyy")

    def test_price_cells_use_thousands_separator_number_format(self):
        tmp, path = generate_layout_workbook()
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        line_font = font_for_style(styles, cell_style(sheet, "E24"))
        self.assertEqual(num_fmt_for_style(styles, cell_style(sheet, "E24")), "4")
        self.assertIsNone(line_font.find(f"{NS_MAIN}b"))
        self.assertIsNone(line_font.find(f"{NS_MAIN}i"))
        self.assertNotEqual(font_color(line_font).get("rgb"), "FFFF0000")

        for ref in ("E27", "E28", "E29"):
            font = font_for_style(styles, cell_style(sheet, ref))
            self.assertEqual(num_fmt_for_style(styles, cell_style(sheet, ref)), "4")
            self.assertIsNotNone(font.find(f"{NS_MAIN}b"))
            self.assertIsNone(font.find(f"{NS_MAIN}i"))
            self.assertNotEqual(font_color(font).get("rgb"), "FFFF0000")

    def test_notes_are_plain_numbered_and_acceptance_text_stays_out_of_notes(self):
        explicit_notes = ["First explicit note", "Second explicit note"]
        tmp, path = generate_layout_workbook({
            "notes_heading": "Note : ",
            "standard_notes": explicit_notes,
        })
        self.addCleanup(tmp.cleanup)

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(zf.read("xl/styles.xml"))

        self.assertEqual(cell_value(sheet, "A33"), "1.00")
        self.assertEqual(cell_value(sheet, "A34"), "2.00")
        self.assertEqual(cell_value(sheet, "B33"), explicit_notes[0])
        self.assertIsNone(font_for_style(styles, cell_style(sheet, "A33")).find(f"{NS_MAIN}i"))
        self.assertIsNone(font_for_style(styles, cell_style(sheet, "A34")).find(f"{NS_MAIN}i"))
        self.assertEqual(cell_value(sheet, "E35"), "")
        self.assertEqual(cell_value(sheet, "E37"), "We accept the quotation amount and the terms")

    def test_brief_can_override_terms_notes_and_company_text(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "quotation.xlsx"
        brief = {
            "company_identity": "Other Company",
            "quote_date": "2026-06-04",
            "project_number": "OC-001",
            "client": {"name": "Sample Client", "attention": "Alex Tan"},
            "project": {"title": "Sample Project"},
            "line_items": [],
            "payment_terms": [
                "50% deposit",
                "50% before handover",
                "Warranty excluded",
                "All cheques should be crossed and made payable to Other Company Pte Ltd",
            ],
            "terms_heading": "Commercial Terms",
            "cheque_payee": "Other Company Pte Ltd",
            "notes_heading": "Editable Notes",
            "standard_notes": ["First editable note", "Second editable note"],
            "company": {
                "name": "Other Company Pte Ltd",
                "header_lines": ["Other Company Pte Ltd", "Dynamic address line", "", "Dynamic bank line"],
                "logo_data_url": "data:image/png;base64,iVBORw0KGgo=",
            },
            "acceptance": {
                "company_name": "Other Company Pte Ltd",
                "text": "Accepted by customer",
                "person_label": "Authorised signer",
                "stamp_label": "Customer stamp",
                "date_label": "Signed date:",
            },
            "signature": {
                "koncept_signatory": "Morgan Lee",
                "koncept_title": "Sales Lead",
                "koncept_date_label": "Company signed date:",
            },
        }
        price = quote.PriceRow(1, "Floor", "m2 needle punch carpet in colour", "sqm", 7, 1.09, 1.5, "")
        line = quote.QuoteLine("Floor", 1, "sqm", "Needle punch carpet", "needle punch carpet in colour", "", price, 10.5, "matched", [])

        quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, [line])

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            drawing = ET.fromstring(zf.read("xl/drawings/drawing1.xml"))
            drawing_rels = zf.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
            media_names = set(zf.namelist())

        self.assertEqual(find_cell_ref(sheet, "Commercial Terms"), "A32")
        self.assertEqual(find_cell_ref(sheet, "50% deposit"), "B33")
        self.assertEqual(find_cell_ref(sheet, "Warranty excluded"), "B35")
        self.assertEqual(find_cell_ref(sheet, "All cheques should be crossed and made payable to Other Company Pte Ltd"), "B36")
        self.assertTrue(find_cell_ref(sheet, "Editable Notes").startswith("A"))
        self.assertTrue(find_cell_ref(sheet, "First editable note").startswith("B"))
        self.assertTrue(find_cell_ref(sheet, "Second editable note").startswith("B"))
        self.assertTrue(find_cell_ref(sheet, "Other Company Pte Ltd").startswith("B"))
        self.assertTrue(find_cell_ref(sheet, "Accepted by customer").startswith("E"))
        self.assertTrue(find_cell_ref(sheet, "Authorised signer").startswith("E"))
        self.assertTrue(find_cell_ref(sheet, "Customer stamp").startswith("E"))
        self.assertTrue(find_cell_ref(sheet, "Company signed date:").startswith("B"))
        self.assertTrue(find_cell_ref(sheet, "Signed date:").startswith("E"))

        paragraphs = [
            "".join(text_node.text or "" for text_node in paragraph.findall(f".//{NS_A}t"))
            for paragraph in drawing.findall(f".//{NS_A}p")
        ]
        self.assertIn("Other Company Pte Ltd", paragraphs)
        self.assertIn("Dynamic address line", paragraphs)
        self.assertEqual(paragraphs[2], "")
        self.assertIn("Dynamic bank line", paragraphs)
        self.assertIn("Target=\"../media/header_logo.png\"", drawing_rels)
        self.assertIn("xl/media/header_logo.png", media_names)

    def test_payment_terms_cheque_instruction_is_not_written_twice(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "quotation.xlsx"
        cheque_instruction = "All cheques should be crossed and made payable to Other Company Pte Ltd"
        brief = {
            "company_identity": "Other Company",
            "quote_date": "2026-06-04",
            "client": {"name": "Sample Client", "attention": "Alex Tan"},
            "project": {"title": "Sample Project"},
            "line_items": [],
            "payment_terms": ["50% deposit", cheque_instruction],
            "terms_heading": "Commercial Terms",
            "cheque_payee": "Other Company Pte Ltd",
            "notes_heading": "Editable Notes",
            "standard_notes": ["First editable note"],
            "company": {"name": "Other Company Pte Ltd"},
            "rich_text": {
                "paymentTerms": (
                    "<div><strong>50%</strong> deposit</div>"
                    "<div>All cheques should be crossed and made payable to <strong>Other Company Pte Ltd</strong></div>"
                ),
            },
        }
        price = quote.PriceRow(1, "Floor", "m2 needle punch carpet in colour", "sqm", 7, 1.09, 1.5, "")
        line = quote.QuoteLine("Floor", 1, "sqm", "Needle punch carpet", "needle punch carpet in colour", "", price, 10.5, "matched", [])

        quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, [line])

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        self.assertEqual(find_cell_refs(sheet, cheque_instruction), ["B34"])
        self.assertEqual(
            cell_inline_runs(sheet, "B34"),
            [
                ("All cheques should be crossed and made payable to ", False, False, False),
                ("Other Company Pte Ltd", True, False, False),
            ],
        )
        self.assertEqual(cell_value(sheet, "B35"), "")
        self.assertEqual(find_cell_ref(sheet, "Editable Notes"), "A36")

    def test_quote_detail_rich_text_runs_are_written_to_layout_output(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "quotation.xlsx"
        brief = {
            "company_identity": "Other Company",
            "quote_date": "2026-06-04",
            "client": {
                "name": "Sample Client",
                "attention": "Alex Tan",
                "address": ["10 Sample Street", "Singapore 000010"],
            },
            "project": {"title": "Sample Project"},
            "line_items": [],
            "payment_terms": ["50% deposit"],
            "terms_heading": "Commercial Terms",
            "cheque_payee": "Other Company Pte Ltd",
            "notes_heading": "Editable Notes",
            "standard_notes": ["First editable note"],
            "company": {
                "name": "Other Company Pte Ltd",
                "header_lines": ["Other Company Pte Ltd", "Dynamic address line"],
            },
            "rich_text": {
                "quoteDate": "<div><strong><em><u>04 June 2026</u></em></strong></div>",
                "clientAddress": "<div><strong>10 Sample</strong> <em>Street</em></div><div><u>Singapore 000010</u></div>",
                "headerDetails": "<div><strong>Other Company Pte Ltd</strong></div><div><em><u>Dynamic address line</u></em></div>",
                "termsHeading": "<div><strong>Commercial Terms</strong></div>",
                "paymentTerms": "<div><strong>50%</strong> <em>deposit</em></div>",
                "standardNotes": "<div>First <u>editable</u> note</div>",
            },
        }
        price = quote.PriceRow(1, "Floor", "m2 needle punch carpet in colour", "sqm", 7, 1.09, 1.5, "")
        line = quote.QuoteLine("Floor", 1, "sqm", "Needle punch carpet", "needle punch carpet in colour", "", price, 10.5, "matched", [])

        quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, [line])

        with zipfile.ZipFile(path) as zf:
            sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
            drawing = ET.fromstring(zf.read("xl/drawings/drawing1.xml"))

        self.assertEqual(
            cell_inline_runs(sheet, "A7"),
            [
                ("10 Sample", True, False, False),
                (" ", False, False, False),
                ("Street", False, True, False),
            ],
        )
        self.assertEqual(cell_inline_runs(sheet, "A8"), [("Singapore 000010", False, False, True)])
        self.assertEqual(cell_inline_runs(sheet, "A16"), [("04 June 2026", True, True, True)])
        self.assertEqual(cell_inline_runs(sheet, find_cell_ref(sheet, "Commercial Terms")), [("Commercial Terms", True, False, False)])
        self.assertEqual(
            cell_inline_runs(sheet, "B33"),
            [
                ("50%", True, False, False),
                (" ", False, False, False),
                ("deposit", False, True, False),
            ],
        )
        note_ref = find_cell_ref(sheet, "First editable note")
        self.assertTrue(note_ref.startswith("B"))
        self.assertEqual(
            cell_inline_runs(sheet, note_ref),
            [
                ("First ", False, False, False),
                ("editable", False, False, True),
                (" note", False, False, False),
            ],
        )
        header_runs = drawing_paragraph_runs(drawing)
        self.assertIn([("Other Company Pte Ltd", True, False, False)], header_runs)
        self.assertIn([("Dynamic address line", False, True, True)], header_runs)

    def test_missing_logo_removes_template_logo_from_output(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "quotation.xlsx"
        brief = {
            "company_identity": "No Logo Company",
            "quote_date": "2026-06-04",
            "project_number": "NO-LOGO-001",
            "client": {"name": "Sample Client", "attention": "Alex Tan"},
            "project": {"title": "No Logo Project"},
            "line_items": [],
            "payment_terms": [],
            "company": {"name": "No Logo Company", "header_lines": ["No Logo Company"]},
        }
        price = quote.PriceRow(1, "Floor", "m2 needle punch carpet in colour", "sqm", 7, 1.09, 1.5, "")
        line = quote.QuoteLine("Floor", 1, "sqm", "Needle punch carpet", "needle punch carpet in colour", "", price, 10.5, "matched", [])

        quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, [line])

        with zipfile.ZipFile(path) as zf:
            drawing = ET.fromstring(zf.read("xl/drawings/drawing1.xml"))
            drawing_rels = zf.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
            media_names = set(zf.namelist())

        self.assertFalse(any(anchor.find(f"{NS_DRAWING}pic") is not None for anchor in drawing.findall(f"{NS_DRAWING}twoCellAnchor")))
        self.assertNotIn("/image", drawing_rels)
        self.assertNotIn("xl/media/image1.jpeg", media_names)

    def test_header_logo_replacement_does_not_hijack_unrelated_picture(self):
        parts = {
            "xl/drawings/drawing1.xml": drawing_with_picture_xml("Product Screenshot", "rId5"),
            "xl/drawings/_rels/drawing1.xml.rels": drawing_rels_xml(("rId5", "../media/product.png")),
            "xl/media/product.png": b"product image",
            "[Content_Types].xml": empty_content_types_xml(),
        }

        quote.replace_header_logo(parts, "data:image/png;base64,iVBORw0KGgo=")

        drawing = ET.fromstring(parts["xl/drawings/drawing1.xml"])
        rels = ET.fromstring(parts["xl/drawings/_rels/drawing1.xml.rels"])
        pics = drawing.findall(f".//{NS_DRAWING}pic")
        names = [
            pic.find(f"{NS_DRAWING}nvPicPr/{NS_DRAWING}cNvPr").attrib.get("name")
            for pic in pics
        ]
        targets = {
            rel.attrib.get("Id"): rel.attrib.get("Target")
            for rel in rels.findall(f"{NS_PACKAGE_REL}Relationship")
        }

        self.assertIn("Product Screenshot", names)
        self.assertIn("Header Logo", names)
        self.assertEqual(targets["rId5"], "../media/product.png")
        self.assertIn("../media/header_logo.png", targets.values())
        self.assertEqual(parts["xl/media/product.png"], b"product image")
        self.assertIn("xl/media/header_logo.png", parts)

    def test_header_logo_replacement_creates_missing_drawing_rels_file(self):
        parts = {
            "xl/drawings/drawing1.xml": empty_drawing_xml(),
            "[Content_Types].xml": empty_content_types_xml(),
        }

        quote.replace_header_logo(parts, "data:image/png;base64,iVBORw0KGgo=")

        self.assertIn("xl/drawings/_rels/drawing1.xml.rels", parts)
        drawing = ET.fromstring(parts["xl/drawings/drawing1.xml"])
        rels = ET.fromstring(parts["xl/drawings/_rels/drawing1.xml.rels"])
        header_pic = next(
            pic
            for pic in drawing.findall(f".//{NS_DRAWING}pic")
            if pic.find(f"{NS_DRAWING}nvPicPr/{NS_DRAWING}cNvPr").attrib.get("name") == "Header Logo"
        )
        header_rel_id = header_pic.find(f".//{NS_A}blip").attrib[f"{NS_REL}embed"]
        header_rel = next(
            rel
            for rel in rels.findall(f"{NS_PACKAGE_REL}Relationship")
            if rel.attrib.get("Id") == header_rel_id
        )

        self.assertEqual(header_rel.attrib["Target"], "../media/header_logo.png")
        self.assertIn("xl/media/header_logo.png", parts)

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
            quote.write_quote_layout_xlsx(KONCEPT_LAYOUT, path, brief, [line])

            with zipfile.ZipFile(path) as zf:
                sheet = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        self.assertEqual(worksheet_formulas(sheet), ["SUM(E22:E26)", "SUM(E27:E28)"])
        self.assertEqual(cell_value(sheet, "A6"), brief["client"]["name"])
        self.assertEqual(cell_value(sheet, "A18"), brief["project"]["title"])
        self.assertEqual(cell_value(sheet, "C22"), line.section)
        self.assertEqual(cell_value(sheet, "E24"), line.display_price)
        self.assertTrue(find_cell_ref(sheet, brief["payment_terms"][0]).startswith("B"))

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
        self.assertIn("pricing_id", rows[0])
        self.assertIn("catalog_description", rows[0])
        self.assertNotIn("template_row", rows[0])
        self.assertNotIn("template_description", rows[0])
        self.assertEqual(data[0], "'@manual")
        self.assertEqual(data[1], "'=WEBSERVICE(\"https://example.test/section\")")
        self.assertEqual(data[2], "'+HYPERLINK(\"https://example.test/line\",\"line\")")
        self.assertEqual(data[3], "'-10+20")


if __name__ == "__main__":
    unittest.main()
