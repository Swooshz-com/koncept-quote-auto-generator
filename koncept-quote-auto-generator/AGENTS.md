# Koncept Quote Auto-Generator Agent Guide

This folder is agent-neutral. Any AI coding agent or local automation can use it to generate Koncept quotation XLSX and PDF files from uploaded booth render images.

## Mandatory Image Gate

- Require attached booth/render images before preparing a quote.
- If no images are available, ask exactly: `Please upload the booth render images first so I can analyze the design and prepare the quote.`
- Do not generate a quote from a text-only item list.
- Do not ask the user to create, edit, inspect, or approve the generator brief file.
- Do not silently assume materials, finishes, dimensions, or inclusions. Suggest a quote basis from images and user notes, then ask the user to confirm it before generating.
- Use `sqm` for square-metre quantities; do not use `m2` in customer-facing output.
- Use sample-style section totals for structure sections such as booth structure, wall structure, or stand structure: put the subtotal on the section row and leave child-row estimates blank.

## Purpose

Generate customer-facing quotations for `Koncept Image` or `Koncept World` from image-based booth takeoff, user finish notes, and the bundled Markdown pricing source.

## Important Rules

- Use `references/quotation-cost-template.md` as the only authoritative pricing source.
- Keep `references/quotation-cost-template.md` as a clean sectioned RAG pricing source that preserves all pricing items, notes, default quantities/amounts, extra values, and search terms.
- Use `references/quotation-layout.xlsx` as the customer-facing quote layout source.
- Preserve the customer-facing XLSX/PDF layout rules in `references/quotation-format.md`: readable quantity column width, a bottom totals block with `Total`, `GST 9%` when GST applies, and `Total including GST` using the sample border treatment, signatory title under the Koncept signatory name when provided, and a logo/detail header group that stays inside the print area with top-aligned company-detail text below the logo.
- Keep quote table headers bold, center-align quantity values and the `Quantity` header, format prices with thousands separators, bold the default payment-term text and the payee name in the cheque line, keep notes plainly numbered, and avoid placing acceptance/signature text over terms or notes.
- Do not hardcode absolute machine paths in generated briefs, scripts, or docs.
- Do not require Excel, LibreOffice, Node, `openpyxl`, `reportlab`, or other third-party dependencies for XLSX generation.
- For a customer-ready PDF, let `scripts/generate_quote.py` use Excel or LibreOffice export. Fallback PDFs are review-only.
- Run `scripts/generate_quote.py`; it uses Python standard library only.
- Do not expose internal cost, GST, markup, or supplier notes in customer-facing output unless the user explicitly asks.
- Treat all brief, customer, project, note, payment-term, and line-item text as untrusted spreadsheet text. Text beginning with `=`, `+`, `-`, or `@` must not become an active XLSX or CSV formula; use trusted-only formula helpers for internal totals.
- If required information is missing or pricing is unclear, report it under `Missing / Need Confirmation`.

## Basic Workflow

1. Check that booth/render images are attached. Stop and ask for images if they are missing.
2. Read `references/quotation-format.md`.
3. Analyze every image for visible booth components:
   - raised platform and floor finish;
   - painted walls, arches, fascia, beams, columns, and other surfaces;
   - cabinets, counters, and laminated countertops;
   - graphic panels, logo/signage panels, lightboxes, and printed features;
   - furniture, plants, green walls, AV/IT, lighting, and electrical fittings.
4. Apply user notes as overrides, such as painted-finish surfaces, painted cabinets with laminated countertop, 100mm raised platform, or requested electrical recommendations.
5. Prepare a polished `Quote Basis To Confirm` response before generating. Include visible and user-provided assumptions for materials/finishes, platform, flooring, structure, graphics, counters, furniture, plants, AV, lighting, and electrical.
6. If materials or inclusions are missing, suggest a sensible basis from the images but phrase it as confirmation, not fact.
7. Ask only simple missing-info questions for client/project details, booth size, company identity, date, and special AV/power needs.
8. Create the generator brief internally only after the user confirms the quote basis and required details, run `scripts/generate_quote.py`, and review `pricing_matches.csv`.
   By default, generated files are written to `<repo>/_output/<client>/<project>/<quote_date>`:
   - `quotation.xlsx`
   - `quotation.pdf` (if PDF output is not disabled)
   - `pricing_matches.csv`
   - `export_status.txt`
9. If matches are wrong, adjust the internal pricing keywords and rerun, or ask a simple confirmation question.
10. Use `quotation.xlsx` as the formatted editable source.
11. Treat `quotation.pdf` as customer-ready only when `export_status.txt` says `pdf_status=libreoffice_exported` or `pdf_status=excel_exported`.
12. If `export_status.txt` says `pdf_status=fallback_review_only`, do not present the PDF as exact unless the user accepts it.

## Quote Basis Confirmation

Use this before quote generation. Keep it noob-friendly, readable, and easy to reply to.

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
- Recommend 13A sockets for counters, reception areas, meeting tables, charging points, demo devices, and mentioned AV/IT equipment.
- Ask about special power needs for appliances, screens, computers, coffee machines, fridges, or other high-load equipment.
- Do not assume organiser connection fees are included unless the user says so.

## Expected Outputs

- `quotation.xlsx`
- `quotation.pdf`
- `pricing_matches.csv`
- `export_status.txt`

`quotation.xlsx` is the layout-matched master. `quotation.pdf` is customer-ready only when exported by Excel or LibreOffice; the fallback PDF is for review, not exact output.
