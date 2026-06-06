# Quotation Format Reference

Use this guide to format Koncept quotation outputs. Past quotation files are reference-only and are not packaged into this skill. The packaged `quotation-layout.xlsx` is a cleaned layout template, not an old customer quote.

## Page Header

- Put `ESTIMATE` near the top-right for Koncept Image-style quotes when space allows.
- Put the client block at top-left:
  - Client company name.
  - Address lines.
  - Country or city line.
- Add `Attention: <name>` and the contact title on the next line if available.
- Add the quote date below the attention block.
- Add `RE: <project title>` before the line-item table.
- Add company details and bank details near the top-right for Koncept Image-style quotes when generating XLSX.
- Treat the formatted XLSX as the customer-ready output. Generate PDF only when explicitly requested, and keep it secondary to the XLSX.

## Company Identities

### Koncept Image

- Legal/payment name in terms: `Koncept Image Pte Ltd` or `Koncept Image Pte Limited`, matching the quote brief.
- Known address: `61 Kaki Bukit Avenue 1, #02-26, Shunli Industrial Park, Singapore 417943`.
- Known telephone: `+65 6817 7477`.
- Known bank detail: United Overseas Bank Limited, 80 Raffles Place, Singapore 048624.

### Koncept World

- Legal/payment name in terms: `Koncept World Pte Ltd` or `Koncept World Private Limited`, matching the quote brief.
- Known address: `61 Kaki Bukit Avenue 1, #02-26, Shun Li Industrial Park, Singapore 417943`.
- Known telephone: `+65 9180 3079`.
- Website/email may appear when provided by the user.

Always ask which identity to use. Do not choose by default.

## Line Item Table

- Columns: `Pos.`, `Quantity`, `Service`, `Estimate`.
- Make `Pos.`, `Quantity`, and `Service` bold.
- Center-align quantity values and the `Quantity` header.
- Currency: usually `SGD`.
- Format all numeric price/estimate cells with thousands separators, for example `1,000.00` and `10,000.00`.
- Use `sqm` for square-metre quantities, not `m2`.
- Keep the `Quantity` column wide enough that entries such as `24 m length` and `36 sqm` do not clip in Excel print-to-PDF output.
- Write customer/brief-sourced values as literal spreadsheet text. Do not turn values beginning with `=`, `+`, `-`, or `@` into active formulas in XLSX or CSV outputs.
- Section rows use numbering such as `1.0`, `2.0`, `3.0`.
- Detail rows use numbering such as `1.1`, `1.2`, `2.1`.
- Section rows may carry a lump-sum amount while child rows have blank estimates.
- Structure sections such as booth structure, wall structure, or stand structure should use the lump-sum section-row style unless the user asks for itemized pricing.
- Detail rows can display numeric price, `FOC`, or `Included`.
- Multi-line service descriptions are acceptable.
- Do not show internal cost, GST, markup, or red remark labels in the customer quote.

## Common Sections

- Floor Design or Flooring.
- Booth Structure or Wall Structure.
- Special Components.
- Furniture and miscellaneous rental.
- AV/IT equipment rental.
- Stand illumination and electrical fittings.
- Graphic work.
- Assembly / Disassembly.
- Transportation.
- Project Management.
- Waste disposal.

Use only sections needed by the quote brief.

## Totals And Notes

- Add `Total` near the end of line items before GST.
- When GST applies, show `GST 9%` under `Total` without `@`.
- Show `Total including GST` as the final total label.
- Keep `Total`, GST, and `Total including GST` in the bottom totals area of the estimate page when the preserved layout has room.
- Add a clean top rule above `Total`, a clean top rule above `Total including GST`, and a stronger bottom rule under `Total including GST` across the label, amount, and currency cells.
- If discount is provided, show:
  - `Total Estimates items ...`
  - `Less goodwill discount`
  - `Total Final Price`
- Add exclusions before terms when provided.
- Add special notes in red styling in XLSX when possible.
- Keep note numbering plain and sequential. Do not make a single note number bold/italic unless every note number uses the same style.
- Keep acceptance/signature text out of the terms and notes body so it cannot overlap long note lines.

## Export Status

- `quotation.xlsx` is the master output.
- PDF generation is disabled by default; `export_status.txt` should report `pdf_status=skipped` and `pdf_mode=none` for normal runs.
- If a PDF is explicitly requested, generate it with `--pdf-mode auto` and treat it as secondary because the XLSX is the formatting source of truth.

## Terms

Common payment terms:

- `70% payment upon confirmation and signing of contract.`
- `30% balance upon handover before show starts`
- `60% payment upon confirmation and signing of contract.`
- `40% balance 14 days after delivery`

Bold the payment-term text in the customer quote. In the cheque instruction line, keep the sentence regular and bold only the payee name, for example `Koncept Image Pte Ltd`.

Common notes:

- The quote excludes application fees to relevant authorities and organiser electrical connection fees unless stated otherwise.
- Design changes during work may delay completion and are at the client's cost.
- Changes after confirmation are treated as additional orders.
- Designs and dimensions are subject to final site verification.
- Quotation must be confirmed minimum 20 working days before the event for production.
- Graphic cost may have a surcharge if files are late.
- Design and artwork of graphics are not included unless stated.
- Cancellation is subject to 75% of the agreement amount.
- Deposits are non-refundable upon cancellation.
- Late payment charge may apply after due date.

## Signature

Include two signature areas:

- Left: Koncept company name, signature line, Francies Cheng or Francis Cheng as specified, title if provided.
- Right: `We accept the quotation amount and the terms`, signature line, `Person in charge`, `Company name & stamp`, `Date:`.
- When a Koncept signatory title or designation is supplied, place it directly below the signatory name.
- Keep the logo and company-detail text together inside the `A:I` print area, not against the right page edge. Place the details below, not beside or behind, the logo, with only a tight normal line gap below the logo. Print the company-detail text on page 1 only; repeated pages should keep the logo/header without the address/bank text chunk. Align company-detail text to the logo's left edge, use 9 pt text, left-align every detail line, keep the specified address and bank line breaks, add a normal 9 pt blank line before `Bank Detail:`, and keep the detail text readable in Excel.
