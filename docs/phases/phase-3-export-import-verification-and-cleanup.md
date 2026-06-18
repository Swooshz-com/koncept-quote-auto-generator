# Phase 3 - Safe Export/Import Verification And Repo Cleanup

## Goal

Verify export/import paths before removing obsolete bundled/demo material.

## Required Order

1. Verify export from the current app.
2. Verify import into the intended local/staging company workspace path.
3. Confirm generated quotes still use the expected profile, pricing reference,
   layout, and customer-facing wording.
4. Migrate tests away from bundled/demo files that are no longer intended to be
   active fixtures.
5. Delete only files proven unused by app behavior and tests.

## Cleanup Rules

- Do not delete bundled files used by active tests without replacing fixtures or
  updating tests in the same PR.
- Do not delete canonical docs, validation scripts, pricing-reference guardrails,
  safety gates, or generated-reference builders.
- Do not delete pricing reference images, templates, layouts, or profile assets
  until import/export verification proves the replacement data path works.
- Do not weaken `scripts/validate_dynamic_pricing_reference_rules.py` or any CI
  guard that keeps pricing matching data-driven.

## Exit Criteria

- Export verification evidence is recorded in the PR.
- Import verification evidence is recorded in the PR.
- Replaced tests point at the new fixture/seed path.
- Deletions are limited, reversible, and explained with rollback notes.

## Phase 3A Inventory - 2026-06-18

Phase 3A verified the Koncept Images workspace import path and classified the
repo-bundled data before deleting anything.

### Workspace Path Verification

- Active local/staging company id is `koncept-images-pte-ltd`.
- Imported quote-company profile exports are saved under
  `QUOTE_DATA_ROOT/koncept-images-pte-ltd/profiles.json`.
- The old placeholder `QUOTE_DATA_ROOT/default/profiles.json` path is not used
  for imported Koncept Images profiles.
- Imported profile `logo_data_url` survives save and load, and generated
  `quotation.xlsx` writes the restored logo to `xl/media/header_logo.png`.
- The verification uses a tiny sanitized PNG data URL in tests, not a real
  exported company profile artifact.

### Cleanup Classification

| Path | Classification | Phase 3A decision |
| --- | --- | --- |
| `workspace-seeds/koncept-images-pte-ltd/workspace.json` | Company/workspace seed bridge | Keep. It is the active migration-shaped workspace identity. |
| `profiles/koncept/profile.json` | Active bundled profile seed | Keep. It still resolves the default profile, quote presets, and layout links until imported profile/template storage replaces it. |
| `profiles/koncept/quotation-layout.xlsx` | Active customer-facing layout template | Keep. XLSX generation still depends on it as the selected profile layout source. |
| `profiles/koncept/layout-rules.json` | Active formatting guardrail | Keep. It preserves the customer-facing output formatting rules. |
| `profiles/koncept/assets/koncept-header-logo.jpeg` | Active bundled profile asset | Keep. Current profile preset logo resolution still depends on it. |
| `pricing-references/koncept-exhibition-quotation/reference.json` | Active bundled pricing-reference metadata | Keep. It is the default pricing reference pack metadata. |
| `pricing-references/koncept-exhibition-quotation/pricing-catalog.json` | Active bundled pricing catalog | Keep. Catalog matching, quote basis repair, pricing review, and generation tests still depend on it. |
| `pricing-references/koncept-exhibition-quotation/pricing-catalog.ai-reference.md` | Active AI/catalog prompt reference | Keep. Prompt and matching behavior still use it. |
| `pricing-references/koncept-exhibition-quotation/pricing-catalog-images/*` | Active pricing-reference visual assets | Keep. They are part of the bundled pricing reference pack and import/export visual-reference coverage. |
| `pricing-references/import-cleanup-rules.json` | Import cleanup rules/guardrail | Keep. Not business seed/demo data. |
| `scripts/build_pricing_catalog.py` | Generated-reference builder | Keep. Not business seed/demo data. |
| `scripts/validate_dynamic_pricing_reference_rules.py` | Pricing matching guardrail | Keep. Required to prevent fixture-specific runtime matching patches. |
| `fixtures/samples/kent-group/sample.json` | Active sanitized sample fixture | Keep. `/api/samples`, PDF intake tests, and Playwright smoke flows still use it. |
| `fixtures/samples/kent-group/kent-group.pdf` | Active sample PDF fixture | Keep. PDF rendering/intake tests and local smoke scripts still use it. |

### Removed In Phase 3A

None. The replaceable default/demo cleanup candidate list is empty for this
slice because every bundled profile, pricing, template, logo, visual, and
sample fixture inspected above is still referenced by app behavior, tests, or
guardrails. Future removal should wait until replacement imported profile,
template, pricing-reference, and sample fixture paths are implemented and
covered by tests.
