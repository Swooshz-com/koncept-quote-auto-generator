# Pricing Catalog Import Preview

## Purpose

Pricing catalog import lets Settings create a reviewed company pricing reference from supplier catalog files without saving the raw upload as the final reference.

## Desired AI-assisted flow

1. An admin opens Settings -> Pricing References -> New / Import.
2. The admin enters a Pricing Reference name and chooses the GST/VAT label and rate.
3. The admin uploads a `.xlsx`, `.csv`, or `.md` pricing catalog.
4. The server extracts bounded content only.
5. AI normalizes messy catalog content into pricing reference rows when configured.
6. The UI renders an editable preview table.
7. The admin edits invalid or uncertain rows.
8. Save Reference stays disabled until required fields are valid.
9. Nothing is saved until the admin confirms Save Reference.

## Supported input types

- `.xlsx` workbook files.
- `.csv` files.
- `.md` Markdown catalog notes.

Normalized `.xlsx` and `.csv` templates can be parsed deterministically. Messy workbook, CSV, and Markdown layouts use AI normalization when an AI provider is configured.

## Safe extraction boundaries

- Upload size is bounded before parsing.
- XLSX zip entries, total uncompressed XML, shared strings, rows, and columns are bounded.
- Extracted AI context is row-oriented and limited to non-empty cells and short text snippets.
- Raw upload content is not saved as the pricing reference and is not returned in public responses.
- Formula-like text beginning with `=`, `+`, `-`, or `@` is neutralized before preview/save.

## Multi-row description and remarks stitching

Real catalogs may split one pricing item across physical rows. Import should detect continuation rows where core pricing fields are blank but description or remark cells contain text. Continuation description text is appended to the previous item with `; `. Continuation remark text is appended to previous remarks with `; `.

The stitcher must preserve casing, all-caps remarks, dimensions, units, and technical wording. It must not merge independent priced rows such as rigging points when those rows include their own unit, cost, markup, or other pricing identity.

For Koncept V1.1 workbooks, bullet or note-style continuation rows belong in remarks, not in the customer-facing pricing reference description.

Import cleanup has two layers:

- Deterministic local rules in `pricing-references/import-cleanup-rules.json` handle safe known cleanup such as typography, word-slash spacing, customer-facing `m2` to `sqm`, and audited typo replacements.
- AI normalization for messy/unrecognized workbooks should infer additional obvious spelling, OCR, spacing, and unit cleanup from that workbook's own repeated terms, nearby rows, section headings, and standard unit notation. AI must not paraphrase, market-polish, simplify, or rename technical catalog descriptions.

After save, the pricing reference text is authoritative. Later quote-basis and output logic must use the saved customer-facing description word for word.

## Saved order and matching metadata

Saved pricing rows carry:

- `category_order`: the source category index, or first-seen category order when the source has no numeric index.
- `item_order`: the source row order.
- `match_terms`: deterministic item-level search terms derived from the saved description, aliases, remarks, and unit.
- `object_families`: broad catalog families such as graphics, display, water, partition, fascia, or beverage_service.

Quote-basis sorting should prefer `category_order` and `item_order` when the selected sorting mode follows the pricing reference. AI matching and repair should use saved `match_terms` and `object_families` from the pricing reference instead of adding customer/sample-specific runtime keyword patches.

## AI normalization schema

AI normalization returns rows with these fields:

- `section`
- `description`
- `unit_hint`
- `internal_cost`
- `markup_multiplier`
- `remarks`
- `aliases`
- `match_terms`
- `object_families`
- `warning` or `status`

AI should preserve short technical catalog rows, include aliases/remarks useful for retrieval, and keep commercial notes in remarks instead of the main customer-facing description. Deterministic import enrichment may regenerate `match_terms` and `object_families` on save, so AI-provided values are suggestions, not authority.

## Editable preview table

The preview table exposes section, description, unit hint, internal cost, markup multiplier, remarks, aliases, and warning/status. Invalid rows remain visible for correction where possible. Save is disabled until required fields are valid and duplicate IDs are handled.

## User confirmation before save

Preview generation is not persistence. The company pricing reference is written only after the admin clicks Save Reference.

## Deferred items

- Full DB-backed persistence.
- Advanced multi-sheet workbook inference.
- Richer human-in-the-loop AI editing UI.
- Full profile CRUD.
- Deployment/auth architecture changes.
