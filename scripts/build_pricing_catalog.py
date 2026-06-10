#!/usr/bin/env python3
"""Build a structured pricing catalog from an uploaded pricing workbook."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import generate_quote as quote


CATALOG_SCHEMA_VERSION = 1

COL_SECTION_NO = 0
COL_DEFAULT_QUANTITY = 1
COL_DESCRIPTION = 2
COL_DEFAULT_ESTIMATE = 5
COL_COST = 7
COL_GST = 8
COL_MARKUP = 9
COL_REMARKS = 11


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a Koncept pricing workbook into a root-level pricing reference catalog.")
    parser.add_argument("--source", required=True, type=Path, help="Pricing source .xlsx file.")
    parser.add_argument("--source-label", help=argparse.SUPPRESS)
    parser.add_argument("--out", required=True, type=Path, help="Output pricing catalog JSON path.")
    parser.add_argument("--ai-reference-md-out", type=Path, help="Optional generated Markdown view for AI catalog reading.")
    return parser.parse_args()


def normalized_text(value: Any) -> str:
    text = str(value or "")
    text = text.replace("\xa0", " ")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    return re.sub(r"\s+", " ", text).strip()


def numeric_value(value: Any) -> float | None:
    if value in (None, ""):
        return None
    parsed = quote.as_float(value, 0.0)
    if parsed == 0.0 and str(value).strip() not in {"0", "0.0", "0.00"}:
        return None
    return parsed


def cell(row: list[Any], index: int) -> Any:
    return row[index] if index < len(row) else None


def text_cell(row: list[Any], index: int) -> str:
    return normalized_text(cell(row, index))


def is_section_row(row: list[Any]) -> bool:
    return numeric_value(cell(row, COL_SECTION_NO)) is not None and bool(text_cell(row, COL_DESCRIPTION)) and numeric_value(cell(row, COL_COST)) is None


def is_price_row(row: list[Any]) -> bool:
    return bool(text_cell(row, COL_DESCRIPTION)) and (numeric_value(cell(row, COL_COST)) or 0.0) > 0


def append_unique(values: list[str], value: str) -> None:
    cleaned = normalized_text(value)
    if not cleaned:
        return
    lowered = {item.lower() for item in values}
    if cleaned.lower() not in lowered:
        values.append(cleaned)


def append_description_part(item: dict[str, Any], value: str) -> None:
    cleaned = normalized_text(value)
    if not cleaned:
        return
    if cleaned.startswith("\u2022") or cleaned.lower().startswith(("note:", "notes:")):
        append_unique(item["remarks"], cleaned)
        return
    append_unique(item["description_parts"], cleaned)


def append_remark(item: dict[str, Any], value: str) -> None:
    append_unique(item["remarks"], value)


def strip_leading_unit(value: str) -> str:
    return normalized_text(
        re.sub(
            r"^(?:m2|sqm|m\.?\s*length|m\.?\s*run|nos\.?|no\.|sets?|lot\.?)\s+(?:of\s+|rental\s+of\s+)?",
            "",
            value,
            flags=re.IGNORECASE,
        )
    )


def alias_candidates(section: str, description_parts: list[str], remarks: list[str], unit_hint: str) -> list[str]:
    aliases: list[str] = []
    source_values = [*description_parts, *remarks]
    for value in source_values:
        append_unique(aliases, value)
        append_unique(aliases, strip_leading_unit(value))
        for part in re.split(r"[;/,]", value):
            append_unique(aliases, strip_leading_unit(part))
    if section:
        for value in description_parts:
            append_unique(aliases, f"{section} {strip_leading_unit(value)}")
    if unit_hint:
        for value in description_parts:
            append_unique(aliases, f"{strip_leading_unit(value)} {unit_hint}")
    return aliases


def catalog_id(section: str, description: str, existing_ids: set[str]) -> str:
    section_slug = quote.slugify_segment(section, "pricing")
    item_slug = quote.slugify_segment(strip_leading_unit(description), "item")
    base = f"{section_slug}.{item_slug}"
    candidate = base
    suffix = 2
    while candidate in existing_ids:
        candidate = f"{base}-{suffix}"
        suffix += 1
    existing_ids.add(candidate)
    return candidate


def finalize_item(item: dict[str, Any], items: list[dict[str, Any]]) -> None:
    if not item:
        return
    description_parts = item.pop("description_parts")
    remarks = item["remarks"]
    description = "; ".join(description_parts)
    unit_hint = quote.infer_unit(description)
    section = item["section"]
    existing_ids = {existing["id"] for existing in items}
    item.update(
        {
            "id": catalog_id(section, description, existing_ids),
            "description": description,
            "unit_hint": unit_hint,
            "aliases": alias_candidates(section, description_parts, remarks, unit_hint),
            "sale_unit_price": round(float(item["internal_cost"]) * float(item["markup_multiplier"]), 2),
        }
    )
    items.append(item)


def item_from_price_row(section: str, row: list[Any]) -> dict[str, Any]:
    default_quantity = numeric_value(cell(row, COL_DEFAULT_QUANTITY))
    default_estimate = numeric_value(cell(row, COL_DEFAULT_ESTIMATE))
    gst_multiplier = numeric_value(cell(row, COL_GST)) or 1.0
    item = {
        "section": section,
        "description_parts": [text_cell(row, COL_DESCRIPTION)],
        "default_quantity": default_quantity,
        "default_quote_amount": default_estimate if default_estimate and default_estimate > 0 else None,
        "internal_cost": numeric_value(cell(row, COL_COST)) or 0.0,
        "gst_multiplier": gst_multiplier,
        "markup_multiplier": numeric_value(cell(row, COL_MARKUP)) or 1.0,
        "remarks": [],
        "extra_values": [],
    }
    append_remark(item, text_cell(row, COL_REMARKS))
    for value in row[COL_REMARKS + 1:]:
        cleaned = normalized_text(value)
        if cleaned:
            item["extra_values"].append(cleaned)
    return item


def build_catalog_from_xlsx(source: Path) -> dict[str, Any]:
    items: list[dict[str, Any]] = []
    rows = quote.read_first_sheet_rows_with_numbers(source)
    current_section = ""
    current_item: dict[str, Any] | None = None

    for _row_number, row in rows:
        if is_section_row(row):
            finalize_item(current_item or {}, items)
            current_item = None
            current_section = text_cell(row, COL_DESCRIPTION)
            continue

        if is_price_row(row):
            finalize_item(current_item or {}, items)
            current_item = item_from_price_row(current_section, row)
            continue

        if current_item is None:
            continue

        description = text_cell(row, COL_DESCRIPTION)
        remark = text_cell(row, COL_REMARKS)
        if description:
            append_description_part(current_item, description)
        if remark:
            append_remark(current_item, remark)
        for value in row[COL_REMARKS + 1:]:
            cleaned = normalized_text(value)
            if cleaned:
                current_item["extra_values"].append(cleaned)

    finalize_item(current_item or {}, items)
    return catalog_payload(items)


def catalog_payload(items: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "schema_version": CATALOG_SCHEMA_VERSION,
        "currency": "SGD",
        "items": items,
    }


def build_catalog(source: Path, source_label: str | None = None) -> dict[str, Any]:
    suffix = source.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return build_catalog_from_xlsx(source)
    raise ValueError(f"Pricing catalog source must be an .xlsx workbook: {source}")


def write_catalog(catalog: dict[str, Any], out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(catalog, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def catalog_to_ai_reference_markdown(catalog: dict[str, Any], catalog_name: str = "pricing-catalog.json") -> str:
    lines = [
        "---",
        "title: Pricing Catalog AI Reference",
        "kind: generated_pricing_ai_reference",
        f"source_catalog: {catalog_name}",
        "---",
        "",
        "# Pricing Catalog AI Reference",
        "",
        "Generated from the structured pricing catalog. Do not edit this file directly.",
        "Use this Markdown for AI readability; use the JSON catalog for pricing logic.",
        "",
    ]
    current_section = None
    for item in catalog.get("items", []):
        section = normalized_text(item.get("section")) or "Unsectioned"
        if section != current_section:
            lines.extend([f"## {section}", ""])
            current_section = section
        lines.extend(
            [
                f"### {item.get('id', '')}",
                "",
                f"- **Item:** {item.get('description', '')}",
                f"- **Unit hint:** {item.get('unit_hint') or 'not specified'}",
            ]
        )
        if item.get("default_quantity") is not None:
            lines.append(f"- **Default quantity:** {item.get('default_quantity')}")
        if item.get("default_quote_amount") is not None:
            lines.append(f"- **Default quote amount:** SGD {item.get('default_quote_amount')}")
        remarks = item.get("remarks") or []
        if remarks:
            lines.append(f"- **Remarks:** {'; '.join(str(value) for value in remarks)}")
        aliases = item.get("aliases") or []
        if aliases:
            lines.append(f"- **Search aliases:** {'; '.join(str(value) for value in aliases[:12])}")
        extra_values = item.get("extra_values") or []
        if extra_values:
            values = "; ".join(str(value) for value in extra_values[:8])
            lines.append(f"- **Extra values:** {values}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def write_ai_reference_markdown(catalog: dict[str, Any], out: Path, catalog_name: str = "pricing-catalog.json") -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(catalog_to_ai_reference_markdown(catalog, catalog_name), encoding="utf-8")


def main() -> None:
    args = parse_args()
    catalog = build_catalog(args.source, args.source_label)
    write_catalog(catalog, args.out)
    print(f"Wrote {args.out}")
    if args.ai_reference_md_out:
        write_ai_reference_markdown(catalog, args.ai_reference_md_out, args.out.name)
        print(f"Wrote {args.ai_reference_md_out}")


if __name__ == "__main__":
    main()
