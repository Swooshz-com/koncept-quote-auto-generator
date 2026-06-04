---
name: koncept-quote-auto-generator
description: Generate company quotation documents for Koncept Image or Koncept World using the bundled `_Quotation Cost Template V1.1.xlsx` as the pricing source. Use when an AI agent or local automation needs to create, revise, price-check, or export a Koncept-style booth quotation from uploaded booth render images plus a plain-English user request. Image attachments are mandatory; if no booth images are shared, ask the user to upload them before preparing the quote.
---

# Koncept Quote Auto-Generator

This folder is usable by any AI coding agent or direct local script. The user should be able to ask casually and upload booth images; the agent handles the takeoff and generator brief internally.

## Core Rules

- Require attached booth/render images before quote generation.
- If no images are available, ask exactly: `Please upload the booth render images first so I can analyze the design and prepare the quote.`
- Do not generate a quote from a text-only item list.
- Do not ask the user to create, edit, inspect, or approve the generator brief file.
- Do not silently assume materials, finishes, dimensions, or inclusions. Suggest a quote basis from images and user notes, then ask the user to confirm it before generating.
- Use `sqm` for square-metre quantities; do not use `m2` in customer-facing output.
- Use sample-style section totals for structure sections such as booth structure, wall structure, or stand structure: put the subtotal on the section row and leave child-row estimates blank.
- Use `_Quotation Cost Template V1.1.xlsx` beside this `SKILL.md` as the only pricing source.
- Use `references/quotation-layout.xlsx` as the customer-facing quote layout source.
- Preserve the customer-facing XLSX/PDF layout rules in `references/quotation-format.md`: the quantity column must be wide enough for values like `24 m length`, GST and Grand Total rows must have clear top/bottom rules, totals should stay near the bottom of the estimate page, the Koncept signatory title should appear below the signatory name when provided, and the company-detail text below the logo must be top-aligned and not cramped.
- Do not hardcode absolute user machine paths.
- Do not require Excel, LibreOffice, Node, `openpyxl`, `reportlab`, or other installed dependencies for XLSX generation.
- For a customer-ready PDF, let `scripts/generate_quote.py` use Excel or LibreOffice export. Fallback PDFs are review-only.
- Use `scripts/generate_quote.py`; it is written for Python standard library only.
- Do not copy internal cost, GST, markup, or supplier notes into the customer-facing quotation unless the user explicitly asks.

## Image-Drop Quote Workflow

1. Check for attached booth/render images. Stop and ask for images if none are available.
2. Read `references/quotation-format.md`.
3. Inspect all images and create an internal visual takeoff from visible booth components:
   - raised platform and floor finish;
   - painted walls, arches, fascia, beams, columns, and other surfaces;
   - cabinets, counters, and laminated countertops;
   - graphic panels, logo/signage panels, lightboxes, and printed features;
   - furniture, plants, green walls, AV/IT, lighting, and electrical fittings.
4. Treat the user's text as instructions or overrides, such as painted-finish surfaces, painted cabinets with laminated countertop, 100mm raised platform, or recommended electrical fittings.
5. Prepare a polished `Quote Basis To Confirm` response before generating. Include visible and user-provided assumptions for:
   - painted, laminate, fabric, acrylic, glass, carpet, vinyl, platform, and other material/finish choices;
   - raised platform height and whether it covers the full booth area;
   - flooring type and booth size or confirmed dimensions;
   - walls, arches, fascia, beams, columns, cabinets, counters, and countertops;
   - graphics, signage, lightboxes, logo panels, and printed features;
   - furniture, plants, green walls, AV/IT, lighting, electrical fittings, and optional items.
6. If materials or inclusions are missing, suggest a sensible quote basis from the images but phrase it as confirmation, not fact.
7. Ask only simple missing-info questions:
   - client name and attention person;
   - event/project name;
   - booth size or confirmed dimensions;
   - quote date and project number if needed;
   - `Koncept Image` or `Koncept World`;
   - AV, screens, appliances, demo devices, or special power needs.
8. Create the generator brief internally only after the user confirms the quote basis and required details, then run:

```bash
python scripts/generate_quote.py --brief path/to/internal-brief-file --allow-ambiguous
```

9. Review `pricing_matches.csv`, `export_status.txt`, the XLSX, and the PDF export status. If pricing is unmatched or visibly wrong, refine the internal keywords and rerun, or ask the user a simple confirmation question.
10. Output files are written by default to `<repo>/_output/<client>/<project>/<quote_date>` as:
    - `quotation.xlsx`
    - `quotation.pdf` (if PDF output is not disabled)
    - `pricing_matches.csv`
    - `export_status.txt`
11. Treat `quotation.pdf` as customer-ready only when `export_status.txt` says `pdf_status=libreoffice_exported` or `pdf_status=excel_exported`.
12. If `export_status.txt` says `pdf_status=fallback_review_only`, treat `quotation.xlsx` as the polished master and do not present the PDF as exact unless the user accepts it.
13. Deliver links only after totals, visible quote text, and PDF export status are checked.

## Quote Basis Confirmation

Use a concise, well-formatted confirmation response before quote generation. Keep it noob-friendly and avoid technical estimator jargon.

Formatting rules:

- Start with the Markdown heading `**Quote Basis To Confirm**`.
- Use bold labels for every category.
- Under each category, use point form only.
- Start bullets with `Include:`, `Confirm:`, `Exclude:`, or `Note:`.
- Keep bullets short, ideally one line each.
- Do not write paragraph-style category descriptions.
- Split missing client/project details into a separate `**Missing Info Needed**` section.
- End with one bold confirmation sentence.
- Do not bury important questions in a long paragraph.

```text
**Quote Basis To Confirm**

**Surfaces / Structures**
- Include: <visible structure items>
- Confirm: <material/finish basis>

**Cabinets / Counters**
- Include: <visible cabinet/counter items>
- Confirm: <paint finish and countertop material>

**Platform / Flooring**
- Include: <platform/flooring items>
- Confirm: <platform height, coverage, and flooring finish>

**Graphics / Signage**
- Include: <visible panels, signs, lightboxes, and printed features>
- Confirm: <any unclear artwork or logo scope>

**Furniture / Plants / AV**
- Include: <visible/requested furniture, plants, green walls, AV/IT>
- Confirm: <any unclear AV or rental items>

**Electrical**
- Include: <visible lights and recommended sockets>
- Confirm: <special power, appliances, AV, or organiser connection fees>

**Missing Info Needed**

- Client name and attention person
- Event/project name
- Booth size or confirmed dimensions
- Quote date/project number, if any
- Koncept Image or Koncept World
- Any AV, appliances, demo devices, or special power needs

**Please confirm or correct the quote basis above before I generate the quotation.**
```

If the user already gave a clear material note, use it in the basis. If they did not, write `Please confirm material/finish for this item` instead of inventing it.

## Electrical Recommendation Guidance

- Count visible downlights, spotlights, and exterior sign lights when clear from the images.
- Recommend 13A sockets for counters, reception areas, meeting tables, charging points, demo devices, and any mentioned AV/IT equipment.
- Ask about special power needs when the images or prompt suggest appliances, screens, computers, coffee machines, fridges, or other high-load equipment.
- Do not assume organiser connection fees are included unless the user says so.

## Output Contract

The generator writes:

- `quotation.xlsx`
- `quotation.pdf`
- `pricing_matches.csv`
- `export_status.txt`

The XLSX is the editable formatted source and should match the preserved quote layout. The PDF is customer-ready only when `export_status.txt` reports `libreoffice_exported` or `excel_exported`; otherwise the script creates a styled fallback that is useful for review but not exact.
