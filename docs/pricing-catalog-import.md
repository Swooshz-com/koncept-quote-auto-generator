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

## AI normalization schema

AI normalization returns rows with these fields:

- `section`
- `description`
- `unit_hint`
- `internal_cost`
- `markup_multiplier`
- `remarks`
- `aliases`
- `warning` or `status`

AI should preserve short technical catalog rows, include aliases/remarks useful for retrieval, and keep commercial notes in remarks instead of the main customer-facing description.

## Editable preview table

The preview table exposes section, description, unit hint, internal cost, markup multiplier, remarks, aliases, and warning/status. Invalid rows remain visible for correction where possible. Save is disabled until required fields are valid and duplicate IDs are handled.

## User confirmation before save

Preview generation is not persistence. The company pricing reference is written only after the admin clicks Save Reference.

## Deferred items

- Advanced multi-sheet workbook inference.
- Full human-in-the-loop AI chat editing.
- DB-backed persistence.
- Company role/session integration.
