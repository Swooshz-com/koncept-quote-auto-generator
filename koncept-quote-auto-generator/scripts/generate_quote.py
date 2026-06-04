#!/usr/bin/env python3
"""Generate Koncept quotation XLSX and PDF files.

This script intentionally uses Python standard library only. It reads the
bundled XLSX cost template through ZIP/XML parsing and fills a preserved quote
layout workbook through ZIP/XML updates so the customer-facing XLSX keeps the
same styling, print setup, drawings, and pagination rules as the reference
quotation.
"""

from __future__ import annotations

import argparse
import copy
import csv
import datetime as dt
import html
import io
import json
import math
import os
import re
import shutil
import subprocess
import textwrap
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


SKILL_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = SKILL_DIR / "_Quotation Cost Template V1.1.xlsx"
DEFAULT_LAYOUT_TEMPLATE = SKILL_DIR / "references" / "quotation-layout.xlsx"
DEFAULT_OUTPUT_ROOT = SKILL_DIR / "_output"
NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_CONTENT_TYPES = "{http://schemas.openxmlformats.org/package/2006/content-types}"
NS_PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
NS_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
XMLNS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XMLNS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
XMLNS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
XMLNS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
XMLNS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
EXPORT_STATUS_CUSTOMER_READY = {"libreoffice_exported", "excel_exported"}

ET.register_namespace("", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("mc", XMLNS_MC)
ET.register_namespace("x14ac", XMLNS_X14AC)
ET.register_namespace("xr", XMLNS_XR)
ET.register_namespace("xr2", XMLNS_XR2)
ET.register_namespace("xr3", XMLNS_XR3)
ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("a16", "http://schemas.microsoft.com/office/drawing/2014/main")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate Koncept quotation XLSX and PDF files.")
    parser.add_argument("--brief", required=True, type=Path, help="Path to quote brief JSON.")
    parser.add_argument(
        "--out",
        type=Path,
        help="Output folder. Defaults to _output/<client>/<project>/<YYYYMMDD> when omitted.",
    )
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Cost template XLSX path.")
    parser.add_argument("--layout-template", type=Path, default=DEFAULT_LAYOUT_TEMPLATE, help="Customer quotation layout XLSX path.")
    parser.add_argument("--pdf-mode", choices=("auto", "styled", "text", "none"), default="auto", help="PDF export mode. auto tries Excel/LibreOffice then styled fallback, styled skips external PDF export, text writes a simple fallback, none skips PDF.")
    parser.add_argument("--allow-ambiguous", action="store_true", help="Use the best pricing match even when multiple rows match.")
    return parser.parse_args()


@dataclass
class PriceRow:
    row_number: int
    section: str
    description: str
    unit_hint: str
    cost: float
    gst_multiplier: float
    markup: float
    remark: str

    @property
    def sale_unit_price(self) -> float:
        return round(self.cost * self.markup, 2)


@dataclass
class QuoteLine:
    section: str
    quantity: float | None
    unit: str
    description: str
    pricing_keyword: str
    display_price: str
    matched_price: PriceRow | None
    amount: float | None
    match_status: str
    match_candidates: list[PriceRow]


def col_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    total = 0
    for ch in letters:
        total = total * 26 + (ord(ch.upper()) - 64)
    return total - 1


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        xml = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(xml)
    strings: list[str] = []
    for si in root.findall(f"{NS_MAIN}si"):
        parts = []
        for t in si.iter(f"{NS_MAIN}t"):
            parts.append(t.text or "")
        strings.append("".join(parts))
    return strings


def cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    value_node = cell.find(f"{NS_MAIN}v")
    inline_node = cell.find(f"{NS_MAIN}is")
    if cell_type == "inlineStr" and inline_node is not None:
        return "".join(t.text or "" for t in inline_node.iter(f"{NS_MAIN}t")).strip()
    if value_node is None:
        return None
    raw = value_node.text or ""
    if cell_type == "s":
        return shared_strings[int(raw)] if raw else ""
    if cell_type == "str":
        return raw
    try:
        number = float(raw)
        return int(number) if number.is_integer() else number
    except ValueError:
        return raw


def read_first_sheet_rows(xlsx_path: Path) -> list[list[Any]]:
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_xml = zf.read("xl/worksheets/sheet1.xml")
    root = ET.fromstring(sheet_xml)
    rows: list[list[Any]] = []
    for row in root.iter(f"{NS_MAIN}row"):
        values: list[Any] = []
        for cell in row.findall(f"{NS_MAIN}c"):
            ref = cell.attrib.get("r", "A1")
            col_index = col_to_index(ref)
            while len(values) <= col_index:
                values.append(None)
            values[col_index] = cell_value(cell, shared_strings)
        rows.append(values)
    return rows


def as_float(value: Any, default: float = 0.0) -> float:
    if value in (None, ""):
        return default
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return default


def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_unit(unit: Any) -> str:
    text = clean_text(unit)
    lower = text.lower()
    if lower in {"m2", "m^2", "sq m", "sq.m", "sq.m.", "square metre", "square meter", "square metres", "square meters"}:
        return "sqm"
    return text


def slugify_segment(value: Any, fallback: str = "item") -> str:
    raw = clean_text(value).lower()
    slug = re.sub(r"[^a-z0-9]+", "-", raw).strip("-")
    return slug or fallback


def resolve_default_output_dir(brief: dict[str, Any], provided_out: Path | None) -> Path:
    if provided_out is not None:
        return provided_out
    client_name = slugify_segment((brief.get("client") or {}).get("name"), "client")
    project_name = slugify_segment((brief.get("project") or {}).get("title"), "project")
    date_part = slugify_segment((brief.get("quote_date") or dt.date.today().isoformat()), "quote")
    return DEFAULT_OUTPUT_ROOT / client_name / project_name / date_part


def infer_unit(description: str) -> str:
    text = description.lower()
    for unit in ("m2", "sqm", "m length", "m run", "nos", "no.", "lot", "sets"):
        if unit in text:
            return normalize_unit(unit)
    return ""


def extract_price_rows(template_path: Path) -> list[PriceRow]:
    rows = read_first_sheet_rows(template_path)
    price_rows: list[PriceRow] = []
    current_section = ""
    for index, row in enumerate(rows, start=1):
        padded = row + [None] * 15
        col_a = padded[0]
        description = clean_text(padded[2])
        cost = as_float(padded[7], 0.0)
        gst = as_float(padded[8], 1.0)
        markup = as_float(padded[9], 1.0)
        remark = clean_text(padded[11])
        if isinstance(col_a, (int, float)) and description and cost == 0:
            current_section = description
            continue
        if description and cost > 0:
            price_rows.append(
                PriceRow(
                    row_number=index,
                    section=current_section,
                    description=description,
                    unit_hint=infer_unit(description),
                    cost=cost,
                    gst_multiplier=gst or 1.0,
                    markup=markup or 1.0,
                    remark=remark,
                )
            )
    return price_rows


def tokens(text: str) -> set[str]:
    stop = {
        "and",
        "or",
        "of",
        "in",
        "at",
        "with",
        "for",
        "the",
        "as",
        "per",
        "ready",
        "construct",
        "not",
        "real",
        "template",
        "item",
        "items",
    }
    return {t for t in re.findall(r"[a-z0-9]+", text.lower()) if len(t) > 1 and t not in stop}


def score_price_row(query: str, row: PriceRow) -> int:
    query_tokens = tokens(query)
    haystack = f"{row.section} {row.description} {row.remark} {row.unit_hint}"
    row_tokens = tokens(haystack)
    overlap = len(query_tokens & row_tokens)
    phrase_bonus = 5 if query.lower() in haystack.lower() else 0
    return overlap + phrase_bonus


def find_price_match(query: str, price_rows: list[PriceRow]) -> tuple[str, PriceRow | None, list[PriceRow]]:
    if not query:
        return "manual-display", None, []
    scored = [(score_price_row(query, row), row) for row in price_rows]
    scored = [(score, row) for score, row in scored if score > 0]
    scored.sort(key=lambda item: (-item[0], item[1].row_number))
    candidates = [row for _, row in scored[:5]]
    if not candidates:
        return "unmatched", None, []
    top_score = scored[0][0]
    tied = [row for score, row in scored if score == top_score]
    if len(tied) > 1:
        return "ambiguous", tied[0], tied[:5]
    return "matched", candidates[0], candidates


REQUIRED_TOP_LEVEL = ("company_identity", "quote_date", "client", "project", "line_items")


def load_brief(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as f:
        return json.load(f)


def validate_brief(brief: dict[str, Any]) -> list[str]:
    missing: list[str] = []
    for key in REQUIRED_TOP_LEVEL:
        if not brief.get(key):
            missing.append(key)
    client = brief.get("client") or {}
    project = brief.get("project") or {}
    if not client.get("name"):
        missing.append("client.name")
    if not client.get("attention"):
        missing.append("client.attention")
    if not project.get("title"):
        missing.append("project.title")
    if brief.get("company_identity") not in {"Koncept Image", "Koncept World"}:
        missing.append("company_identity must be Koncept Image or Koncept World")
    return missing


def prepare_lines(brief: dict[str, Any], price_rows: list[PriceRow], allow_ambiguous: bool) -> list[QuoteLine]:
    prepared: list[QuoteLine] = []
    for item in brief.get("line_items", []):
        display_price = str(item.get("display_price") or "")
        query = clean_text(item.get("pricing_keyword") or item.get("description") or "")
        status, match, candidates = find_price_match(query, price_rows)
        quantity = item.get("quantity")
        quantity_num = as_float(quantity, 0.0) if quantity not in (None, "") else None
        amount: float | None = None
        if display_price:
            status = "manual-display"
        elif status == "matched" or (status == "ambiguous" and allow_ambiguous):
            amount = round((quantity_num or 0.0) * (match.sale_unit_price if match else 0.0))
            if status == "ambiguous" and allow_ambiguous:
                status = "matched-from-ambiguous"
        prepared.append(
            QuoteLine(
                section=clean_text(item.get("section")),
                quantity=quantity_num,
                unit=normalize_unit(item.get("unit")),
                description=clean_text(item.get("description")),
                pricing_keyword=query,
                display_price=display_price,
                matched_price=match if amount is not None else None,
                amount=amount,
                match_status=status,
                match_candidates=candidates,
            )
        )
    return prepared


def confirmation_issues(missing: list[str], lines: list[QuoteLine]) -> list[str]:
    issues = [f"Missing required field: {field}" for field in missing]
    for line in lines:
        if line.match_status == "unmatched":
            issues.append(f"Unmatched pricing: {line.description} / keyword `{line.pricing_keyword}`")
        if line.match_status == "ambiguous":
            options = "; ".join(
                f"row {row.row_number}: {row.section} - {row.description} ({row.sale_unit_price:.2f})"
                for row in line.match_candidates[:3]
            )
            issues.append(f"Ambiguous pricing: {line.description}. Candidate matches: {options}")
    return issues


def print_missing_confirmation(issues: list[str]) -> None:
    print("Missing / Need Confirmation")
    for issue in issues:
        print(f"- {issue}")


def excel_col(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def xlsx_cell(row: int, col: int, value: Any) -> str:
    ref = f"{excel_col(col)}{row}"
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    safe = html.escape(str(value))
    return f'<c r="{ref}" t="inlineStr"><is><t>{safe}</t></is></c>'


def sheet_xml(rows: list[list[Any]]) -> str:
    row_xml = []
    for r_index, row in enumerate(rows, start=1):
        cells = "".join(xlsx_cell(r_index, c_index, value) for c_index, value in enumerate(row))
        row_xml.append(f'<row r="{r_index}">{cells}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData>'
        + "".join(row_xml)
        + '</sheetData></worksheet>'
    )


def write_minimal_xlsx(path: Path, rows: list[list[Any]]) -> None:
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Quotation" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''
    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml(rows))


def pdf_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def write_text_pdf(path: Path, title: str, lines: list[str]) -> None:
    width, height = 595, 842
    margin_x, y_start, line_height = 42, 800, 14
    pages = [lines[i:i + 50] for i in range(0, len(lines), 50)] or [[]]
    objects: list[bytes] = []
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    page_refs = " ".join(f"{3 + i * 2} 0 R" for i in range(len(pages)))
    objects.append(f"<< /Type /Pages /Kids [{page_refs}] /Count {len(pages)} >>".encode("ascii"))
    for page_index, page_lines in enumerate(pages):
        content_obj = 4 + page_index * 2
        stream_lines = ["BT", "/F1 10 Tf", f"{margin_x} {y_start} Td"]
        stream_lines.append(f"({pdf_escape(title)}) Tj")
        stream_lines.append(f"0 -{line_height * 2} Td")
        for line in page_lines:
            stream_lines.append(f"({pdf_escape(line[:110])}) Tj")
            stream_lines.append(f"0 -{line_height} Td")
        stream_lines.append("ET")
        stream = "\n".join(stream_lines).encode("latin-1", errors="replace")
        objects.append(f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Resources << /Font << /F1 {3 + len(pages) * 2} 0 R >> >> /Contents {content_obj} 0 R >>".encode("ascii"))
        objects.append(b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream")
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    chunks = [b"%PDF-1.4\n"]
    offsets = [0]
    for index, obj in enumerate(objects, start=1):
        offsets.append(sum(len(chunk) for chunk in chunks))
        chunks.append(f"{index} 0 obj\n".encode("ascii") + obj + b"\nendobj\n")
    xref_offset = sum(len(chunk) for chunk in chunks)
    chunks.append(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
    for offset in offsets[1:]:
        chunks.append(f"{offset:010d} 00000 n \n".encode("ascii"))
    chunks.append(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii"))
    path.write_bytes(b"".join(chunks))


def money(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"{value:,.2f}"
    return clean_text(value)


def spreadsheet_safe_text(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    stripped = value.lstrip()
    if stripped and stripped[0] in {"=", "+", "-", "@"}:
        return "'" + value
    return value


def is_draft_or_placeholder_note(note: str) -> bool:
    lowered = clean_text(note).lower()
    return any(word in lowered for word in ("draft", "placeholder"))


def customer_notes(brief: dict[str, Any]) -> list[str]:
    return [
        note
        for note in (clean_text(note) for note in brief.get("notes", []))
        if note and not is_draft_or_placeholder_note(note)
    ]


def quote_gst_multiplier(lines: list[QuoteLine]) -> float:
    multipliers = [
        line.matched_price.gst_multiplier
        for line in lines
        if line.matched_price and line.matched_price.gst_multiplier > 1
    ]
    return max(set(multipliers), key=multipliers.count) if multipliers else 1.0


def quote_gst_rate(lines: list[QuoteLine]) -> float:
    return max(quote_gst_multiplier(lines) - 1.0, 0.0)


def gst_label(lines: list[QuoteLine]) -> str:
    rate = quote_gst_rate(lines)
    percent = rate * 100
    if abs(percent - round(percent)) < 0.001:
        percent_text = str(int(round(percent)))
    else:
        percent_text = f"{percent:.2f}".rstrip("0").rstrip(".")
    return f"GST @ {percent_text}%"


def quote_subtotal(entries: list[dict[str, Any]]) -> float:
    total = 0.0
    for entry in entries:
        amount = entry.get("amount")
        if isinstance(amount, (int, float)) and not isinstance(amount, bool):
            total += float(amount)
    return total


def quantity_text(line: QuoteLine) -> str:
    if line.quantity is None:
        return ""
    if float(line.quantity).is_integer():
        qty = str(int(line.quantity))
    else:
        qty = str(line.quantity)
    return f"{qty} {line.unit}".strip()


def build_quote_rows(brief: dict[str, Any], lines: list[QuoteLine]) -> list[list[Any]]:
    client = brief["client"]
    project = brief["project"]
    currency = brief.get("currency", "SGD")
    rows: list[list[Any]] = [
        ["", "", "", "", "ESTIMATE"],
        [client.get("name", "")],
        *[[line] for line in client.get("address", [])],
        [f"Attention: {client.get('attention', '')}"],
        [client.get("title", "")],
        [brief.get("quote_date", "")],
        [f"RE: {project.get('title', '')}"],
        [],
        ["Pos.", "Quantity", "Service", "Estimate"],
        ["", "", "", currency],
    ]
    entries = render_quote_entries(lines, brief)
    for entry in entries:
        if entry["kind"] == "section":
            rows.append([entry["number"], "", entry["section"], money(entry.get("amount"))])
            if entry.get("coverage"):
                rows.append(["", "", "", entry["coverage"]])
            continue
        rows.append([entry["number"], entry["quantity"], " ".join(entry["description_lines"]), money(entry.get("amount"))])
    discount = as_float(brief.get("discount"), 0.0)
    subtotal = max(quote_subtotal(entries) - discount, 0.0)
    gst_amount = round(subtotal * quote_gst_rate(lines)) if quote_gst_rate(lines) else 0
    final_total = subtotal + gst_amount
    rows.extend([[], ["", "", "Grand Total", money(final_total), currency]])
    if discount:
        rows.insert(-1, ["", "", "Less goodwill discount", money(discount), currency])
    if gst_amount:
        rows.insert(-1, ["", "", gst_label(lines), money(gst_amount), currency])
    rows.extend([[], ["Terms & Conditions :"]])
    for idx, term in enumerate(brief.get("payment_terms", []), start=1):
        rows.append([idx, term])
    rows.extend([
        ["Note :"],
        [1, "The above contract does not include application fees to any relevant authorities and electrical connection fees unless stated otherwise."],
        [2, "Any changes in design during the progress of work will delay completion schedule and it shall be deemed at the cost of the Client."],
        [3, "Any changes agreed upon after the confirmation of contract or during the work in progress shall be deemed as Additional Orders."],
        [4, "All designs and dimensions are subject to final site verification."],
        [5, "For production purpose, quotation must be confirmed minimum 20 working days before date of event."],
        [],
        [brief.get("company_identity", "")],
        [],
        ["_____________________________", "", "_____________________________________"],
        [brief.get("signature", {}).get("koncept_signatory", "Francies Cheng"), "", "Person in charge"],
        [brief.get("signature", {}).get("koncept_title", ""), "", "Company name & stamp"],
        ["", "", "Date:"],
    ])
    return rows


def build_pdf_lines(rows: list[list[Any]]) -> list[str]:
    result = []
    for row in rows:
        text = "  ".join(str(value) for value in row if value not in (None, ""))
        if text:
            wrapped = textwrap.wrap(text, width=105) or [text]
            result.extend(wrapped)
        else:
            result.append("")
    return result


def cell_ref(row: int, col: int) -> str:
    return f"{excel_col(col - 1)}{row}"


def parse_cell_ref(ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)([0-9]+)", ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return int(match.group(2)), col_to_index(match.group(1)) + 1


def row_sort_key(row: ET.Element) -> int:
    return int(row.attrib.get("r", "0"))


def cell_sort_key(cell: ET.Element) -> int:
    return parse_cell_ref(cell.attrib.get("r", "A1"))[1]


def get_or_create_row(sheet_data: ET.Element, row_number: int) -> ET.Element:
    for row in sheet_data.findall(f"{NS_MAIN}row"):
        if int(row.attrib.get("r", "0")) == row_number:
            return row
    row = ET.Element(f"{NS_MAIN}row", {"r": str(row_number)})
    rows = sheet_data.findall(f"{NS_MAIN}row")
    insert_at = len(rows)
    for index, existing in enumerate(rows):
        if row_sort_key(existing) > row_number:
            insert_at = index
            break
    sheet_data.insert(insert_at, row)
    return row


def get_or_create_cell(row: ET.Element, row_number: int, col_number: int, style: str | None = None) -> ET.Element:
    ref = cell_ref(row_number, col_number)
    for cell in row.findall(f"{NS_MAIN}c"):
        if cell.attrib.get("r") == ref:
            if style is not None:
                cell.attrib["s"] = style
            return cell
    cell_attrib = {"r": ref}
    if style is not None:
        cell_attrib["s"] = style
    cell = ET.Element(f"{NS_MAIN}c", cell_attrib)
    cells = row.findall(f"{NS_MAIN}c")
    insert_at = len(cells)
    for index, existing in enumerate(cells):
        if cell_sort_key(existing) > col_number:
            insert_at = index
            break
    row.insert(insert_at, cell)
    return cell


def clear_cell(cell: ET.Element) -> None:
    for child in list(cell):
        if child.tag in {f"{NS_MAIN}v", f"{NS_MAIN}is", f"{NS_MAIN}f"}:
            cell.remove(child)
    cell.attrib.pop("t", None)


def set_ooxml_cell(root: ET.Element, row_number: int, col_number: int, value: Any, style: str | None = None) -> None:
    sheet_data = root.find(f"{NS_MAIN}sheetData")
    if sheet_data is None:
        raise ValueError("Layout workbook is missing sheetData.")
    row = get_or_create_row(sheet_data, row_number)
    cell = get_or_create_cell(row, row_number, col_number, style)
    clear_cell(cell)
    if value in (None, ""):
        return
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = ET.SubElement(cell, f"{NS_MAIN}v")
        number.text = str(value)
        return
    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, f"{NS_MAIN}is")
    text = ET.SubElement(inline, f"{NS_MAIN}t")
    text.text = str(value)
    if str(value) != str(value).strip():
        text.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"


def set_ooxml_formula(root: ET.Element, row_number: int, col_number: int, formula: str, style: str | None = None) -> None:
    sheet_data = root.find(f"{NS_MAIN}sheetData")
    if sheet_data is None:
        raise ValueError("Layout workbook is missing sheetData.")
    row = get_or_create_row(sheet_data, row_number)
    cell = get_or_create_cell(row, row_number, col_number, style)
    clear_cell(cell)
    formula_node = ET.SubElement(cell, f"{NS_MAIN}f")
    formula_node.text = formula[1:] if formula.startswith("=") else formula


def set_ooxml_column_width(root: ET.Element, col_number: int, width: float) -> None:
    cols = root.find(f"{NS_MAIN}cols")
    if cols is None:
        cols = ET.Element(f"{NS_MAIN}cols")
        root.insert(0, cols)

    for col in cols.findall(f"{NS_MAIN}col"):
        min_col = int(col.attrib.get("min", "0"))
        max_col = int(col.attrib.get("max", "0"))
        if min_col == col_number and max_col == col_number:
            col.attrib["width"] = str(width)
            col.attrib["customWidth"] = "1"
            return

    cols.append(
        ET.Element(
            f"{NS_MAIN}col",
            {
                "min": str(col_number),
                "max": str(col_number),
                "width": str(width),
                "customWidth": "1",
            },
        )
    )


def clear_ooxml_range(root: ET.Element, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    sheet_data = root.find(f"{NS_MAIN}sheetData")
    if sheet_data is None:
        return
    for row in sheet_data.findall(f"{NS_MAIN}row"):
        row_number = int(row.attrib.get("r", "0"))
        if not min_row <= row_number <= max_row:
            continue
        for cell in row.findall(f"{NS_MAIN}c"):
            _, col_number = parse_cell_ref(cell.attrib.get("r", "A1"))
            if min_col <= col_number <= max_col:
                clear_cell(cell)


def strip_stale_workbook_parts(parts: dict[str, bytes]) -> None:
    parts.pop("xl/sharedStrings.xml", None)
    parts.pop("xl/calcChain.xml", None)
    if "[Content_Types].xml" in parts:
        root = ET.fromstring(parts["[Content_Types].xml"])
        for child in list(root):
            part_name = child.attrib.get("PartName")
            if part_name in {"/xl/sharedStrings.xml", "/xl/calcChain.xml"}:
                root.remove(child)
        parts["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if "xl/_rels/workbook.xml.rels" in parts:
        root = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
        for child in list(root):
            target = child.attrib.get("Target")
            rel_type = child.attrib.get("Type", "")
            if (
                target in {"sharedStrings.xml", "calcChain.xml"}
                or rel_type.endswith("/sharedStrings")
                or rel_type.endswith("/calcChain")
            ):
                root.remove(child)
        parts["xl/_rels/workbook.xml.rels"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def append_border(styles_root: ET.Element, top: str | None = None, bottom: str | None = None) -> int:
    borders = styles_root.find(f"{NS_MAIN}borders")
    if borders is None:
        borders = ET.SubElement(styles_root, f"{NS_MAIN}borders")

    border = ET.Element(f"{NS_MAIN}border")
    ET.SubElement(border, f"{NS_MAIN}left")
    ET.SubElement(border, f"{NS_MAIN}right")
    top_node = ET.SubElement(border, f"{NS_MAIN}top")
    if top:
        top_node.attrib["style"] = top
        ET.SubElement(top_node, f"{NS_MAIN}color", {"auto": "1"})
    bottom_node = ET.SubElement(border, f"{NS_MAIN}bottom")
    if bottom:
        bottom_node.attrib["style"] = bottom
        ET.SubElement(bottom_node, f"{NS_MAIN}color", {"auto": "1"})
    ET.SubElement(border, f"{NS_MAIN}diagonal")

    borders.append(border)
    borders.attrib["count"] = str(len(borders))
    return len(borders) - 1


def clone_cell_style(
    styles_root: ET.Element,
    base_style: str,
    *,
    border_id: int | None = None,
    font_id: str | None = None,
    num_fmt_id: str | None = None,
    horizontal: str | None = None,
    vertical: str | None = None,
) -> str:
    cell_xfs = styles_root.find(f"{NS_MAIN}cellXfs")
    if cell_xfs is None:
        raise ValueError("Layout workbook is missing cellXfs styles.")

    style = copy.deepcopy(cell_xfs[int(base_style)])
    if border_id is not None:
        style.attrib["borderId"] = str(border_id)
        style.attrib["applyBorder"] = "1"
    if font_id is not None:
        style.attrib["fontId"] = font_id
        style.attrib["applyFont"] = "1"
    if num_fmt_id is not None:
        style.attrib["numFmtId"] = num_fmt_id
        style.attrib["applyNumberFormat"] = "1"
    if horizontal is not None or vertical is not None:
        alignment = style.find(f"{NS_MAIN}alignment")
        if alignment is None:
            alignment = ET.SubElement(style, f"{NS_MAIN}alignment")
        if horizontal is not None:
            alignment.attrib["horizontal"] = horizontal
        if vertical is not None:
            alignment.attrib["vertical"] = vertical
        style.attrib["applyAlignment"] = "1"
    cell_xfs.append(style)
    cell_xfs.attrib["count"] = str(len(cell_xfs))
    return str(len(cell_xfs) - 1)


def add_quote_layout_styles(parts: dict[str, bytes]) -> dict[str, str]:
    styles_root = ET.fromstring(parts["xl/styles.xml"])
    gst_border = append_border(styles_root, top="thin")
    grand_border = append_border(styles_root, top="thin", bottom="double")
    bold_header_font = "13"
    style_ids = {
        "header_pos": clone_cell_style(styles_root, "23", font_id=bold_header_font),
        "header_quantity": clone_cell_style(styles_root, "24", font_id=bold_header_font, horizontal="center", vertical="center"),
        "header_service": clone_cell_style(styles_root, "21", font_id=bold_header_font, horizontal="left", vertical="center"),
        "price_amount": clone_cell_style(styles_root, "96", num_fmt_id="4"),
        "gst_label": clone_cell_style(styles_root, "34", border_id=gst_border),
        "gst_amount": clone_cell_style(styles_root, "96", border_id=gst_border, num_fmt_id="4"),
        "gst_currency": clone_cell_style(styles_root, "84", border_id=gst_border),
        "grand_label": clone_cell_style(styles_root, "34", border_id=grand_border),
        "grand_amount": clone_cell_style(styles_root, "96", border_id=grand_border, num_fmt_id="4"),
        "grand_currency": clone_cell_style(styles_root, "84", border_id=grand_border),
    }
    parts["xl/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
    return style_ids


def update_drawing_project_number(xml: bytes, project_number: str) -> bytes:
    project_number = clean_text(project_number)
    if not project_number:
        return xml
    text = xml.decode("utf-8")
    replacement = html.escape(project_number, quote=False)
    updated, count = re.subn(
        r"(<a:t>\s*Project No:\s*</a:t>.*?<a:t>)(.*?)(</a:t>)",
        rf"\g<1>{replacement}\g<3>",
        text,
        count=1,
        flags=re.DOTALL,
    )
    return updated.encode("utf-8") if count else xml


def update_repeating_header_drawing(xml: bytes, project_number: str) -> bytes:
    root = ET.fromstring(xml)
    anchors = root.findall(f"{NS_DRAWING}twoCellAnchor")
    text_anchor = next((anchor for anchor in anchors if anchor.find(f"{NS_DRAWING}sp") is not None), None)
    logo_anchor = next((anchor for anchor in anchors if anchor.find(f"{NS_DRAWING}pic") is not None), None)
    if text_anchor is None:
        return xml

    def update_marker(anchor: ET.Element, marker: str, values: dict[str, str]) -> None:
        marker_node = anchor.find(f"{NS_DRAWING}{marker}")
        if marker_node is None:
            return
        for tag, value in values.items():
            node = marker_node.find(f"{NS_DRAWING}{tag}")
            if node is not None:
                node.text = value

    if logo_anchor is not None:
        update_marker(
            logo_anchor,
            "from",
            {"col": "7", "colOff": "0", "row": "1", "rowOff": "0"},
        )
        update_marker(
            logo_anchor,
            "to",
            {"col": "8", "colOff": "1720000", "row": "2", "rowOff": "415000"},
        )
        pic = logo_anchor.find(f"{NS_DRAWING}pic")
        pic_pr = pic.find(f"{NS_DRAWING}spPr") if pic is not None else None
        pic_xfrm = pic_pr.find(f"{NS_A}xfrm") if pic_pr is not None else None
        logo_off = pic_xfrm.find(f"{NS_A}off") if pic_xfrm is not None else None
        if logo_off is not None:
            logo_off.attrib["x"] = "4550000"
            logo_off.attrib["y"] = "260000"
        logo_ext = pic_xfrm.find(f"{NS_A}ext") if pic_xfrm is not None else None
        if logo_ext is not None:
            logo_ext.attrib["cx"] = "2970000"
            logo_ext.attrib["cy"] = "635000"

    update_marker(text_anchor, "from", {"col": "7", "colOff": "0", "row": "5", "rowOff": "0"})
    update_marker(text_anchor, "to", {"col": "9", "colOff": "200000", "row": "15", "rowOff": "0"})

    sp = text_anchor.find(f"{NS_DRAWING}sp")
    tx_body = sp.find(f"{NS_DRAWING}txBody") if sp is not None else None
    if tx_body is None:
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    body_pr = tx_body.find(f"{NS_A}bodyPr")
    if body_pr is not None:
        body_pr.attrib["vertOverflow"] = "overflow"
        body_pr.attrib["wrap"] = "square"
        body_pr.attrib["anchor"] = "t"
        body_pr.attrib["anchorCtr"] = "0"
        body_pr.attrib["lIns"] = "0"
        body_pr.attrib["rIns"] = "0"
        body_pr.attrib["tIns"] = "0"
        body_pr.attrib["bIns"] = "0"

    for child in list(tx_body):
        if child.tag == f"{NS_A}p":
            tx_body.remove(child)

    lines = [
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
    ]
    if project_number:
        lines.extend(["", f"Project No: {project_number}"])

    for line in lines:
        paragraph = ET.SubElement(tx_body, f"{NS_A}p")
        paragraph_props = ET.SubElement(paragraph, f"{NS_A}pPr")
        paragraph_props.attrib["algn"] = "l"
        run = ET.SubElement(paragraph, f"{NS_A}r")
        run_props = ET.SubElement(run, f"{NS_A}rPr")
        run_props.attrib.update({"lang": "en-US", "sz": "900", "b": "0", "i": "0", "baseline": "0"})
        ET.SubElement(run_props, f"{NS_A}latin").attrib["typeface"] = "+mn-lt"
        ET.SubElement(run_props, f"{NS_A}ea").attrib["typeface"] = "+mn-ea"
        ET.SubElement(run_props, f"{NS_A}cs").attrib["typeface"] = "+mn-cs"
        text = ET.SubElement(run, f"{NS_A}t")
        text.text = line

    sp_pr = sp.find(f"{NS_DRAWING}spPr") if sp is not None else None
    xfrm = sp_pr.find(f"{NS_A}xfrm") if sp_pr is not None else None
    off = xfrm.find(f"{NS_A}off") if xfrm is not None else None
    if off is not None:
        off.attrib["x"] = "4550000"
        off.attrib["y"] = "1350000"
    ext = xfrm.find(f"{NS_A}ext") if xfrm is not None else None
    if ext is not None:
        ext.attrib["cx"] = "3350000"
        ext.attrib["cy"] = "3300000"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def update_print_titles(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    updated = re.sub(
        r"(<definedName[^>]*name=\"_xlnm\.Print_Titles\"[^>]*>[^<]*!\$1:\$)[0-9]+(</definedName>)",
        r"\g<1>5\2",
        text,
        count=1,
    )
    return updated.encode("utf-8")


def excel_date_serial(value: str) -> float | str:
    try:
        parsed = dt.date.fromisoformat(value)
    except ValueError:
        return value
    epoch = dt.date(1899, 12, 30)
    return float((parsed - epoch).days)


def wrapped_description(description: str, width: int = 58) -> list[str]:
    return textwrap.wrap(description, width=width) or [description]


def amount_value(line: QuoteLine) -> str | float | None:
    if line.display_price:
        return line.display_price
    return line.amount


def line_amount_value(line: QuoteLine) -> float | None:
    if line.amount is not None:
        return line.amount
    if line.display_price:
        parsed = as_float(line.display_price, 0.0)
        return parsed if parsed else None
    return None


def next_quote_row(row_number: int) -> int:
    if row_number == 53:
        return 55
    return row_number


def grouped_section_names(brief: dict[str, Any], lines: list[QuoteLine]) -> set[str]:
    configured = brief.get("section_pricing") or {}
    grouped = {
        clean_text(section)
        for section, mode in configured.items()
        if clean_text(mode).lower() in {"section_total", "section-total", "lump_sum", "lump-sum"}
    }
    itemized = {
        clean_text(section)
        for section, mode in configured.items()
        if clean_text(mode).lower() in {"itemized", "line_items", "line-items"}
    }
    for line in lines:
        section = clean_text(line.section)
        section_lower = section.lower()
        if section in itemized:
            continue
        if any(keyword in section_lower for keyword in ("booth structure", "wall structure", "stand structure")):
            grouped.add(section)
    return grouped


def render_quote_entries(lines: list[QuoteLine], brief: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    grouped_sections = grouped_section_names(brief or {}, lines)
    current_section = None
    section_number = 0
    detail_number = 0
    active_section_entry: dict[str, Any] | None = None
    for line in lines:
        if line.section != current_section:
            current_section = line.section
            section_number += 1
            detail_number = 0
            section_grouped = clean_text(current_section) in grouped_sections
            active_section_entry = {
                "kind": "section",
                "number": f"{section_number}.0",
                "section": current_section,
                "amount": 0.0 if section_grouped else None,
                "remark": "",
                "section_grouped": section_grouped,
            }
            entries.append(active_section_entry)
        detail_number += 1
        detail_amount = amount_value(line)
        if active_section_entry and active_section_entry.get("section_grouped"):
            active_section_entry["amount"] = round(float(active_section_entry["amount"] or 0.0) + float(line_amount_value(line) or 0.0), 2)
            detail_amount = None
        entries.append({
            "kind": "item",
            "number": f"{section_number}.{detail_number}",
            "quantity": quantity_text(line),
            "description_lines": wrapped_description(line.description),
            "amount": detail_amount,
        })
    for index, entry in enumerate(entries):
        if entry["kind"] != "section" or not entry.get("section_grouped"):
            continue
        covered_numbers = []
        for follower in entries[index + 1:]:
            if follower["kind"] == "section":
                break
            covered_numbers.append(follower["number"])
        if covered_numbers:
            first, last = covered_numbers[0], covered_numbers[-1]
            entry["coverage"] = f"Covers line items {first} to {last}" if first != last else f"Covers line item {first}"
    return entries


def write_table_header(root: ET.Element, row_number: int, currency_row: int | None = None, styles: dict[str, str] | None = None) -> None:
    styles = styles or {}
    set_ooxml_cell(root, row_number, 1, "Pos.", styles.get("header_pos", "23"))
    set_ooxml_cell(root, row_number, 2, "Quantity", styles.get("header_quantity", "24"))
    set_ooxml_cell(root, row_number, 3, "Service", styles.get("header_service", "21"))
    set_ooxml_cell(root, row_number, 5, "Estimate", "87")
    if currency_row is not None:
        set_ooxml_cell(root, currency_row, 5, "SGD", "95")


def serialize_excel_worksheet(root: ET.Element) -> bytes:
    xml = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")
    worksheet_start_index = xml.find("<worksheet")
    worksheet_tag_end = xml.find(">", worksheet_start_index)
    if worksheet_tag_end == -1:
        return xml.encode("utf-8")

    required_declarations = {
        "xmlns:mc": XMLNS_MC,
        "xmlns:x14ac": XMLNS_X14AC,
        "xmlns:xr": XMLNS_XR,
        "xmlns:xr2": XMLNS_XR2,
        "xmlns:xr3": XMLNS_XR3,
    }
    insertions = []
    worksheet_start = xml[worksheet_start_index:worksheet_tag_end]
    for name, uri in required_declarations.items():
        if f"{name}=" not in worksheet_start:
            insertions.append(f' {name}="{uri}"')
    if insertions:
        xml = xml[:worksheet_tag_end] + "".join(insertions) + xml[worksheet_tag_end:]
    return xml.encode("utf-8")


def write_quote_layout_xlsx(layout_template: Path, path: Path, brief: dict[str, Any], lines: list[QuoteLine]) -> None:
    if not layout_template.exists():
        raise FileNotFoundError(f"Quotation layout template not found: {layout_template}")

    with zipfile.ZipFile(layout_template) as zf:
        parts = {name: zf.read(name) for name in zf.namelist()}

    layout_styles = add_quote_layout_styles(parts)
    root = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    clear_ooxml_range(root, 1, 300, 1, 100)
    set_ooxml_column_width(root, 2, 14.25)
    set_ooxml_column_width(root, 3, 45.5)
    price_style = layout_styles["price_amount"]

    client = brief["client"]
    project = brief["project"]
    currency = brief.get("currency", "SGD")

    set_ooxml_cell(root, 6, 1, client.get("name", ""), "12")
    for offset, line in enumerate((client.get("address") or [])[:4], start=7):
        set_ooxml_cell(root, offset, 1, line, "93")
    set_ooxml_cell(root, 12, 1, f"Attention: {client.get('attention', '')}".strip(), "26")
    set_ooxml_cell(root, 13, 2, client.get("title", ""), "24")
    set_ooxml_cell(root, 16, 1, excel_date_serial(str(brief.get("quote_date", ""))), "101")
    set_ooxml_cell(root, 18, 1, f"RE: {project.get('title', '')}", "26")

    write_table_header(root, 20, 21, layout_styles)
    set_ooxml_cell(root, 21, 5, currency, "95")

    entries = render_quote_entries(lines, brief)
    row_number = 22
    for entry in entries:
        row_number = next_quote_row(row_number)
        if row_number >= 53 and row_number < 55:
            row_number = 55
        if row_number > 92:
            raise ValueError("Quote has too many rows for the preserved layout. Split or shorten line item descriptions.")

        if entry["kind"] == "section":
            set_ooxml_cell(root, row_number, 1, entry["number"], "28")
            set_ooxml_cell(root, row_number, 3, entry["section"], "18")
            if entry.get("amount") is not None:
                set_ooxml_cell(root, row_number, 5, entry["amount"], price_style)
            if entry.get("coverage"):
                set_ooxml_cell(root, row_number + 1, 5, entry["coverage"], "5")
            if entry.get("remark"):
                set_ooxml_cell(root, row_number, 6, entry["remark"], "94")
            row_number += 2
            continue

        description_lines = entry["description_lines"]
        set_ooxml_cell(root, row_number, 1, entry["number"], "70")
        set_ooxml_cell(root, row_number, 2, entry["quantity"], "13")
        set_ooxml_cell(root, row_number, 3, description_lines[0], "5")
        set_ooxml_cell(root, row_number, 5, entry["amount"], price_style)
        for extra in description_lines[1:]:
            row_number += 1
            if row_number == 53:
                row_number = 55
            if row_number > 92:
                raise ValueError("Quote has too many rows for the preserved layout. Split or shorten line item descriptions.")
            set_ooxml_cell(root, row_number, 3, extra, "5")
        row_number += 2

    write_table_header(root, 53, styles=layout_styles)
    gst_rate = quote_gst_rate(lines)
    if gst_rate:
        set_ooxml_cell(root, 93, 4, gst_label(lines), layout_styles["gst_label"])
        set_ooxml_formula(root, 93, 5, f"ROUND(SUM(E22:E92)*{gst_rate:.6f},0)", layout_styles["gst_amount"])
        set_ooxml_cell(root, 93, 6, currency, layout_styles["gst_currency"])
    set_ooxml_cell(root, 94, 4, "Grand Total", layout_styles["grand_label"])
    set_ooxml_formula(root, 94, 5, "SUM(E22:E93)", layout_styles["grand_amount"])
    set_ooxml_cell(root, 94, 6, currency, layout_styles["grand_currency"])

    payment_terms = brief.get("payment_terms") or [
        "80% payment upon confirmation and signing of contract.",
        "20% balance 7 days after delivery.",
    ]
    set_ooxml_cell(root, 99, 1, "Terms & Conditions :", "37")
    for index, term in enumerate(payment_terms[:2], start=1):
        set_ooxml_cell(root, 99 + index, 1, f"{index:.2f}", "40")
        set_ooxml_cell(root, 99 + index, 2, term, "41")
    set_ooxml_cell(root, 102, 2, "All cheques should be crossed and made payable to Koncept Image Pte Ltd", "41")

    standard_notes = [
        "The above contract does not include application fees to any relevant authorities and electrical connection fees.",
        "Any changes in design during the progress of work will delay completion schedule and it shall be deemed at the cost of the Client.",
        "Any changes agreed upon after the confirmation of contract or during the work in progress shall be deemed as Additional Orders.",
        "All designs and dimensions are subject to final site verification.",
        "For production purpose, quotation must be confirmed minimum 20 working days before date of event",
        "20% surcharge will be implied on the graphic cost, if the graphic files are not received latest by five working days before build up date.",
        "Design and Artwork of the graphics are not included in this contract.",
        "Cancellation of agreement is subject to 75% of the agreement amount.",
        "All deposit are non-refundable upon of cancellation of agreement.",
        "The client is obliged to adhere to the above payment terms & conditions of  works.",
        "All payment and/or additional charges shall be settled upon the agreed term of payment schedules.",
        "Late payment charge of 1.5% per month will be charge after the due date.",
    ]
    set_ooxml_cell(root, 103, 1, "Note : ", "37")
    for index, note in enumerate(standard_notes, start=1):
        target_row = 103 + index
        set_ooxml_cell(root, target_row, 1, f"{index:.2f}", "40")
        set_ooxml_cell(root, target_row, 2, note, "41")

    company_name = f"{brief.get('company_identity', 'Koncept Image')} Pte Ltd"
    set_ooxml_cell(root, 117, 2, company_name, "2")
    set_ooxml_cell(root, 117, 5, "We accept the quotation amount and the terms", "2")
    set_ooxml_cell(root, 121, 2, "_____________________________", "33")
    set_ooxml_cell(root, 121, 5, "_____________________________________", "33")
    set_ooxml_cell(root, 122, 2, brief.get("signature", {}).get("koncept_signatory", "Francies Cheng"), "33")
    set_ooxml_cell(root, 122, 5, "Person in charge", "33")
    set_ooxml_cell(root, 123, 2, brief.get("signature", {}).get("koncept_title", ""), "33")
    set_ooxml_cell(root, 123, 5, "Company name & stamp", "33")
    set_ooxml_cell(root, 124, 5, "Date:", "33")

    parts["xl/worksheets/sheet1.xml"] = serialize_excel_worksheet(root)
    if "xl/drawings/drawing1.xml" in parts:
        project_number = clean_text(brief.get("project_number") or project.get("number") or "")
        drawing_xml = update_drawing_project_number(parts["xl/drawings/drawing1.xml"], project_number)
        parts["xl/drawings/drawing1.xml"] = update_repeating_header_drawing(drawing_xml, project_number)
    if "xl/workbook.xml" in parts:
        parts["xl/workbook.xml"] = update_print_titles(parts["xl/workbook.xml"])
    strip_stale_workbook_parts(parts)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in parts.items():
            zf.writestr(name, content)


def powershell_literal(path: Path) -> str:
    return str(path.resolve()).replace("'", "''")


def powershell_export_script(xlsx_path: Path, pdf_path: Path) -> str:
    xlsx_export_path = xlsx_path.resolve()
    pdf_export_path = pdf_path.resolve()
    return f"""
$ErrorActionPreference = 'Stop'
$xlsxPath = '{powershell_literal(xlsx_export_path)}'
$pdfPath = '{powershell_literal(pdf_export_path)}'
$repairedWorkbookPath = [System.IO.Path]::Combine(
  [System.IO.Path]::GetDirectoryName($xlsxPath),
  ([System.IO.Path]::GetFileNameWithoutExtension($xlsxPath) + '.excel-repaired.xlsx')
)
if (Test-Path -LiteralPath $repairedWorkbookPath) {{
  Remove-Item -LiteralPath $repairedWorkbookPath -Force
}}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
  # CorruptLoad 1 lets Excel repair stale template metadata before print export.
  $workbook = $excel.Workbooks.Open($xlsxPath, 0, $false, 5, '', '', $true, 1, '', $false, $false, $null, $false, $true, 1)
  try {{
    $workbook.SaveAs($repairedWorkbookPath, 51)
    $workbook.ExportAsFixedFormat(0, $pdfPath)
  }} finally {{
    $workbook.Close($false)
  }}
}} finally {{
  $excel.Quit()
}}
if (Test-Path -LiteralPath $repairedWorkbookPath) {{
  Move-Item -LiteralPath $repairedWorkbookPath -Destination $xlsxPath -Force
}}
"""


def powershell_pdf_export(xlsx_path: Path, pdf_path: Path) -> str | None:
    if os.name != "nt" or shutil.which("powershell") is None:
        return None
    script = powershell_export_script(xlsx_path, pdf_path)
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        timeout=120,
    )
    return "excel_exported" if result.returncode == 0 and pdf_path.exists() else None


def libreoffice_candidates() -> list[str]:
    candidates = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    if os.name == "nt":
        candidates.extend([
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ])
    seen: set[str] = set()
    result: list[str] = []
    for candidate in candidates:
        if not candidate:
            continue
        path = str(candidate)
        key = path.lower()
        if key not in seen and Path(path).exists():
            seen.add(key)
            result.append(path)
    return result


def libreoffice_pdf_export(xlsx_path: Path, pdf_path: Path) -> str | None:
    executables = libreoffice_candidates()
    if not executables:
        return None
    result = subprocess.run(
        [
            executables[0],
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_path.parent),
            str(xlsx_path),
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
        timeout=120,
    )
    converted = pdf_path.parent / f"{xlsx_path.stem}.pdf"
    if result.returncode == 0 and converted.exists():
        if converted != pdf_path:
            converted.replace(pdf_path)
        return "libreoffice_exported"
    return None


def export_layout_pdf(xlsx_path: Path, pdf_path: Path) -> str | None:
    return libreoffice_pdf_export(xlsx_path, pdf_path) or powershell_pdf_export(xlsx_path, pdf_path)


def pdf_date_text(value: str) -> str:
    try:
        parsed = dt.date.fromisoformat(value)
    except ValueError:
        return value
    return parsed.strftime("%d %B %Y")


def display_amount(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return money(float(value))
    return str(value)


def build_pdf_cell_map(brief: dict[str, Any], lines: list[QuoteLine]) -> dict[tuple[int, int], Any]:
    client = brief["client"]
    project = brief["project"]
    currency = brief.get("currency", "SGD")
    cells: dict[tuple[int, int], Any] = {
        (6, 1): client.get("name", ""),
        (12, 1): f"Attention: {client.get('attention', '')}".strip(),
        (13, 2): client.get("title", ""),
        (16, 1): pdf_date_text(str(brief.get("quote_date", ""))),
        (18, 1): f"RE: {project.get('title', '')}",
        (20, 1): "Pos.",
        (20, 2): "Quantity",
        (20, 3): "Service",
        (20, 5): "Estimate",
        (21, 5): currency,
        (53, 1): "Pos.",
        (53, 2): "Quantity",
        (53, 3): "Service",
        (53, 5): "Estimate",
        (93, 4): gst_label(lines) if quote_gst_rate(lines) else "",
        (94, 4): "Grand Total",
        (94, 6): currency,
        (99, 1): "Terms & Conditions :",
        (102, 2): "All cheques should be crossed and made payable to Koncept Image Pte Ltd",
        (103, 1): "Note : ",
        (106, 5): "We accept the quotation amount and the terms",
        (117, 2): f"{brief.get('company_identity', 'Koncept Image')} Pte Ltd",
        (121, 2): "_____________________________",
        (121, 5): "_____________________________________",
        (122, 2): brief.get("signature", {}).get("koncept_signatory", "Francies Cheng"),
        (122, 5): "Person in charge",
        (123, 5): "Company name & stamp",
        (124, 5): "Date:",
    }
    for offset, address_line in enumerate((client.get("address") or [])[:4], start=7):
        cells[(offset, 1)] = address_line

    entries = render_quote_entries(lines, brief)
    row_number = 22
    for entry in entries:
        row_number = next_quote_row(row_number)
        if row_number >= 53 and row_number < 55:
            row_number = 55
        if row_number > 92:
            break
        if entry["kind"] == "section":
            cells[(row_number, 1)] = entry["number"]
            cells[(row_number, 3)] = entry["section"]
            if entry.get("amount") is not None:
                cells[(row_number, 5)] = entry["amount"]
            if entry.get("coverage"):
                cells[(row_number + 1, 5)] = entry["coverage"]
            if entry.get("remark"):
                cells[(row_number, 6)] = entry["remark"]
            row_number += 2
            continue
        cells[(row_number, 1)] = entry["number"]
        cells[(row_number, 2)] = entry["quantity"]
        cells[(row_number, 3)] = entry["description_lines"][0]
        cells[(row_number, 5)] = entry["amount"]
        for extra in entry["description_lines"][1:]:
            row_number += 1
            if row_number == 53:
                row_number = 55
            if row_number > 92:
                break
            cells[(row_number, 3)] = extra
        row_number += 2

    subtotal = quote_subtotal(entries)
    gst_amount = round(subtotal * quote_gst_rate(lines)) if quote_gst_rate(lines) else 0
    if gst_amount:
        cells[(93, 5)] = gst_amount
        cells[(93, 6)] = currency
    cells[(94, 5)] = subtotal + gst_amount
    payment_terms = brief.get("payment_terms") or [
        "80% payment upon confirmation and signing of contract.",
        "20% balance 7 days after delivery.",
    ]
    for index, term in enumerate(payment_terms[:2], start=1):
        cells[(99 + index, 1)] = f"{index:.2f}"
        cells[(99 + index, 2)] = term

    standard_notes = [
        "The above contract does not include application fees to any relevant authorities and electrical connection fees.",
        "Any changes in design during the progress of work will delay completion schedule and it shall be deemed at the cost of the Client.",
        "Any changes agreed upon after the confirmation of contract or during the work in progress shall be deemed as Additional Orders.",
        "All designs and dimensions are subject to final site verification.",
        "For production purpose, quotation must be confirmed minimum 20 working days before date of event",
        "20% surcharge will be implied on the graphic cost, if the graphic files are not received latest by five working days before build up date.",
        "Design and Artwork of the graphics are not included in this contract.",
        "Cancellation of agreement is subject to 75% of the agreement amount.",
        "All deposit are non-refundable upon of cancellation of agreement.",
        "The client is obliged to adhere to the above payment terms & conditions of  works.",
        "All payment and/or additional charges shall be settled upon the agreed term of payment schedules.",
        "Late payment charge of 1.5% per month will be charge after the due date.",
    ]
    for index, note in enumerate(standard_notes, start=1):
        target_row = 103 + index
        if target_row == 114:
            cells[(target_row, 1)] = "11.00"
            cells[(target_row, 2)] = note
        elif target_row == 115:
            cells[(target_row, 2)] = note
        else:
            cells[(target_row, 1)] = f"{index:.2f}"
            cells[(target_row, 2)] = note
    return cells


def extract_layout_logo(layout_template: Path) -> io.BytesIO | None:
    try:
        with zipfile.ZipFile(layout_template) as zf:
            media_name = next((name for name in zf.namelist() if name.startswith("xl/media/")), None)
            if media_name is None:
                return None
            return io.BytesIO(zf.read(media_name))
    except (OSError, KeyError, zipfile.BadZipFile):
        return None


def company_header_lines() -> list[str]:
    return [
        "Koncept Image Pte Limited",
        "61 Kaki Bukit Ave 1, #02-26, Shunli Industrial Park",
        "Singapore 417943  Tel: +65 6817 7477",
        "",
        "Bank Details: United Overseas Bank Limited",
        "Account: 335-3020-445  Swift Code: UOVBSGSG",
    ]


def write_styled_pdf_fallback(path: Path, brief: dict[str, Any], lines: list[QuoteLine], layout_template: Path) -> bool:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas
    except Exception:
        return False

    cells = build_pdf_cell_map(brief, lines)
    c = canvas.Canvas(str(path), pagesize=A4)
    width, height = A4
    page_ranges = [(1, 52), (53, 95), (96, 124)]
    x_positions = {1: 58, 2: 104, 3: 162, 4: 398, 5: 506, 6: 548}
    row_height = 13.0
    logo = extract_layout_logo(layout_template)

    for page_number, (start_row, end_row) in enumerate(page_ranges, start=1):
        top_y = height - (135 if page_number > 1 else 58)
        if logo is not None:
            logo.seek(0)
            c.drawImage(ImageReader(logo), width - 152, height - 54, width=124, height=25, preserveAspectRatio=True, mask="auto")
            c.setFont("Helvetica", 5.5)
            detail_y = height - 80
            for text in company_header_lines():
                c.drawRightString(width - 28, detail_y, text)
                detail_y -= 7

        for row_number in range(start_row, end_row + 1):
            y = top_y - ((row_number - start_row) * row_height)
            row_values = {col: cells.get((row_number, col)) for col in x_positions}
            if all(value in (None, "") for value in row_values.values()):
                continue
            is_section = row_values.get(1) not in (None, "") and row_values.get(2) in (None, "") and row_values.get(3) not in (None, "") and row_number not in {20, 53}
            is_bold = row_number in {1, 6, 12, 18, 20, 53, 94, 99, 103, 117} or is_section
            c.setFont("Helvetica-Bold" if is_bold else "Helvetica", 10 if row_number < 96 else 8.4)
            if row_number == 94:
                c.setLineWidth(1.4)
                c.line(284, y + 6.5, width - 28, y + 6.5)
            for col_number, value in row_values.items():
                if value in (None, ""):
                    continue
                if page_number == 3 and col_number == 5 and row_number < 117:
                    continue
                text = display_amount(value) if col_number == 5 else str(value)
                x = x_positions[col_number]
                if row_number in {93, 94}:
                    if col_number == 4:
                        c.drawRightString(386, y, text)
                    elif col_number == 5:
                        c.drawCentredString(462, y, text)
                    elif col_number == 6:
                        c.drawRightString(548, y, text)
                    continue
                if col_number == 5 and text in {"Estimate", "SGD"}:
                    c.drawCentredString(x, y, text)
                    continue
                if col_number == 5 and isinstance(value, str):
                    c.drawString(438, y, text[:35])
                    continue
                if page_number == 3 and col_number == 5 and row_number >= 121:
                    c.drawString(345, y, text)
                    continue
                if col_number == 5:
                    c.drawCentredString(x, y, text)
                elif col_number == 6:
                    c.drawRightString(x, y, text)
                else:
                    c.drawString(x, y, text[:78])
        c.setFont("Helvetica", 8)
        c.drawCentredString(width / 2, 24, f"Page {page_number} of {len(page_ranges)}")
        c.showPage()
    c.save()
    return path.exists()


def write_match_csv(path: Path, lines: list[QuoteLine]) -> None:
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["status", "section", "description", "keyword", "template_row", "template_description", "unit_price", "amount"])
        for line in lines:
            match = line.matched_price
            writer.writerow([spreadsheet_safe_text(value) for value in [
                line.match_status,
                line.section,
                line.description,
                line.pricing_keyword,
                match.row_number if match else "",
                match.description if match else "",
                f"{match.sale_unit_price:.2f}" if match else "",
                money(line.amount),
            ]])


def write_export_status(path: Path, status: str, pdf_mode: str) -> None:
    readiness = "customer_ready" if status in EXPORT_STATUS_CUSTOMER_READY else "review_only"
    path.write_text(
        "\n".join([
            f"pdf_status={status}",
            f"pdf_readiness={readiness}",
            f"pdf_mode={pdf_mode}",
        ])
        + "\n",
        encoding="utf-8",
    )


def main() -> int:
    args = parse_args()
    brief = load_brief(args.brief)
    out_dir = resolve_default_output_dir(brief, args.out)
    missing = validate_brief(brief)
    price_rows = extract_price_rows(args.template)
    lines = prepare_lines(brief, price_rows, args.allow_ambiguous)
    issues = confirmation_issues(missing, lines)
    out_dir.mkdir(parents=True, exist_ok=True)
    write_match_csv(out_dir / "pricing_matches.csv", lines)
    if issues:
        print_missing_confirmation(issues)
        return 2
    rows = build_quote_rows(brief, lines)
    xlsx_path = out_dir / "quotation.xlsx"
    pdf_path = out_dir / "quotation.pdf"
    if args.layout_template.exists():
        write_quote_layout_xlsx(args.layout_template, xlsx_path, brief, lines)
    else:
        print(f"Quotation layout template not found, writing minimal XLSX fallback: {args.layout_template}")
        write_minimal_xlsx(xlsx_path, rows)
    pdf_status = "skipped"
    if args.pdf_mode == "auto":
        pdf_status = export_layout_pdf(xlsx_path, pdf_path) or ""
        if not pdf_status:
            print("Excel/LibreOffice PDF export unavailable, writing styled PDF fallback.")
            if not write_styled_pdf_fallback(pdf_path, brief, lines, args.layout_template):
                print("Styled PDF fallback unavailable, writing text PDF fallback.")
                write_text_pdf(pdf_path, f"Quotation - {brief['project']['title']}", build_pdf_lines(rows))
            pdf_status = "fallback_review_only"
    elif args.pdf_mode == "styled":
        if not write_styled_pdf_fallback(pdf_path, brief, lines, args.layout_template):
            print("Styled PDF fallback unavailable, writing text PDF fallback.")
            write_text_pdf(pdf_path, f"Quotation - {brief['project']['title']}", build_pdf_lines(rows))
        pdf_status = "fallback_review_only"
    elif args.pdf_mode == "text":
        write_text_pdf(pdf_path, f"Quotation - {brief['project']['title']}", build_pdf_lines(rows))
        pdf_status = "fallback_review_only"
    write_export_status(out_dir / "export_status.txt", pdf_status, args.pdf_mode)
    print(f"Wrote {out_dir / 'quotation.xlsx'}")
    if args.pdf_mode != "none":
        print(f"Wrote {out_dir / 'quotation.pdf'}")
    print(f"Wrote {out_dir / 'pricing_matches.csv'}")
    print(f"PDF export status: {pdf_status}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
