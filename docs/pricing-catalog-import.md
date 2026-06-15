# Pricing Catalog Import Preview

## Purpose

Pricing catalog import lets Settings create a reviewed company pricing reference from supplier catalog files without saving the raw upload as the final reference.

## Desired AI-assisted flow

1. An admin opens Settings -> Pricing References -> New / Import.
2. The admin enters a Pricing Reference name and chooses the GST/VAT label and rate.
3. The admin uploads a `.xlsx`, `.csv`, or `.md` pricing catalog.
4. The server extracts bounded content only.
5. AI normalizes messy catalog content into pricing reference rows when needed.
6. AI enriches every parsed row with source-backed matching metadata.
7. The UI renders an editable preview table.
8. The admin edits invalid or uncertain rows.
9. Save Reference stays disabled until required fields and AI metadata enrichment are valid.
10. Nothing is saved until the admin confirms Save Reference.

## Supported input types

- `.xlsx` workbook files.
- `.csv` files.
- `.md` Markdown catalog notes.

Normalized `.xlsx` and `.csv` templates can be parsed deterministically for row data, but every user-facing import preview must complete AI metadata enrichment before Save Reference is enabled. Messy workbook, CSV, and Markdown layouts use AI normalization first, then the same mandatory metadata enrichment step.

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

AI provider routing:

- Pricing import normalization and metadata enrichment use `AI_PRICING_IMPORT_PROVIDER` when set.
- When `DEEPSEEK_API_KEY` is present and no explicit provider override is set, DeepSeek is tried first with `DEEPSEEK_MODEL`.
- OpenAI remains the fallback provider and uses `OPENAI_BASIS_LINE_MODEL`.

## Saved order and matching metadata

Saved pricing rows carry:

- `category_order`: the source category index, or first-seen category order when the source has no numeric index.
- `item_order`: the source row order.
- `match_terms`: item-level search terms. Deterministic parsing adds only literal terms from the saved description, aliases, remarks, section, and unit; mandatory AI metadata enrichment adds additional source-backed terms derived from that uploaded catalog.
- `object_families`: source-backed family labels supplied by mandatory AI metadata enrichment or edited reference data. Deterministic parsing preserves this field when provided but does not invent catalog families in Python.

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

AI should preserve short technical catalog rows, include aliases/remarks useful for retrieval, and keep commercial notes in remarks instead of the main customer-facing description. AI should derive `match_terms` and `object_families` from the uploaded catalog's own wording, nearby rows, remarks, headers, and product context. Deterministic save-time enrichment may append literal `match_terms`, but it must not replace mandatory AI metadata, overwrite AI-provided values, or generate hardcoded semantic families.

## Editable preview table

The preview table exposes section, description, unit hint, internal cost, markup multiplier, remarks, aliases, and warning/status. Invalid rows remain visible for correction where possible. Save is disabled until required fields are valid, duplicate IDs are handled, and AI metadata enrichment succeeds.

## User confirmation before save

Preview generation is not persistence. The company pricing reference is written only after the admin clicks Save Reference.

## Deferred items

- Full DB-backed persistence.
- Advanced multi-sheet workbook inference.
- Richer human-in-the-loop AI editing UI.
- Full profile CRUD.
- Deployment/auth architecture changes.
