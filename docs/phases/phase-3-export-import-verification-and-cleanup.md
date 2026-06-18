# Phase 3 - Safe Export/Import Verification And Repo Cleanup

## Goal

Verify export/import paths before removing obsolete bundled/demo material.

## Current Status Update - 2026-06-18

Later cleanup supersedes the earlier "keep" decisions for tracked pricing
references. The repo no longer stores the legacy
`pricing-references/koncept-exhibition-quotation` pack, and local/private
pricing-reference staging belongs under ignored `_pricing-references/` or
runtime/company workspace storage. Repo-safe coverage now uses the synthetic
workspace seed pricing-reference pack.

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
| `_pricing-references/import-cleanup-rules.json` | Optional ignored local import cleanup aid | Runtime/import-specific spelling cleanup belongs under ignored local pricing-reference storage when needed. |
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

## Phase 3B Workspace-Owned Dependency Resolution - 2026-06-18

Phase 3B moves cleanup forward by letting the Koncept Images workspace seed own
the active local/staging runtime dependencies instead of relying on discovered
bundled defaults as the only resolution path.

### Runtime Dependency Model

`workspace-seeds/koncept-images-pte-ltd/workspace.json` now declares
`runtime_dependencies` for:

- active quote-company profile: workspace profile store
  `QUOTE_DATA_ROOT/koncept-images-pte-ltd/profiles.json`, preferred id
  `koncept-images-pte-ltd`
- active logo: the active quote-company profile's `logo_data_url`
- active quotation layout/template: bundled profile `koncept` for this phase
- active layout rules: bundled profile `koncept` for this phase
- active pricing reference: bundled pricing reference
  `koncept-exhibition-quotation` for this phase

This keeps the local/staging model migration-shaped: the workspace records the
company id, profile store, selected quote-company profile, selected layout
source, selected layout-rules source, and selected pricing-reference source that
can later become platform company records.

### Runtime Resolution

- Missing payload `profile_id` now resolves from the workspace's active
  quotation layout dependency before falling back to the repo default profile.
- Missing payload `pricing_reference_id` now resolves from the workspace's
  active pricing-reference dependency before falling back to repo defaults.
- Generated XLSX jobs use the workspace-resolved pricing catalog and layout
  template when payload ids are absent.
- Saved/imported workspace quote-company profiles can fill missing quote-company
  defaults, including `logo_data_url`, before generation validation.
- Bundled quote-company profile presets are exposed only as explicit fallback
  metadata and are not silently applied to generation payloads.

### Bundled Files Still Kept

The same bundled files from Phase 3A remain because they are now fallback/test
fixtures or active workspace-declared dependencies:

- `profiles/koncept/profile.json`
- `profiles/koncept/quotation-layout.xlsx`
- `profiles/koncept/layout-rules.json`
- `profiles/koncept/assets/koncept-header-logo.jpeg`
- `pricing-references/koncept-exhibition-quotation/reference.json`
- `pricing-references/koncept-exhibition-quotation/pricing-catalog.json`
- `pricing-references/koncept-exhibition-quotation/pricing-catalog.ai-reference.md`
- `pricing-references/koncept-exhibition-quotation/pricing-catalog-images/*`
- `fixtures/samples/kent-group/sample.json`
- `fixtures/samples/kent-group/kent-group.pdf`
- `scripts/build_pricing_catalog.py`
- `scripts/validate_dynamic_pricing_reference_rules.py`

### Next Deletion Gate

Do not delete bundled default/demo packs until workspace-owned replacement paths
cover every active runtime and test dependency:

1. imported workspace quote-company profile exists and is selected without
   relying on bundled preset details
2. workspace-owned layout/template and layout-rules source exists outside the
   bundled `profiles/koncept` pack
3. workspace-owned pricing-reference pack exists outside the bundled
   `pricing-references/koncept-exhibition-quotation` pack
4. sample/PDF smoke flows either use sanitized workspace-owned fixtures or are
   replaced by smaller deterministic test fixtures
5. full local tests and CI prove no active app path still requires the old
   bundled files

Phase 3B does not add hosting, deployment, auth, Supabase, Stripe, billing,
OIDC completion, production exposure, or a credit ledger.

## Phase 3C Workspace-Owned Template/Pricing Packs - 2026-06-18

Phase 3C makes the Koncept Images local/staging runtime prefer workspace-owned
seed asset packs for the active quotation template, layout rules, and pricing
reference. Bundled Koncept packs remain in the repo, but they are now explicit
fallback/test fixtures rather than the primary runtime dependency.

### Workspace-Owned Packs

`workspace-seeds/koncept-images-pte-ltd/workspace.json` now declares an
`asset_packs` inventory and `runtime_dependencies` for:

- active quote-company profile: workspace profile store
  `QUOTE_DATA_ROOT/koncept-images-pte-ltd/profiles.json`, preferred id
  `koncept-images-pte-ltd`
- active logo: imported quote-company profile `logo_data_url`
- active quotation layout/template: workspace seed profile pack
  `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/koncept-workspace-template/`
- active layout rules: workspace seed profile pack
  `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/koncept-workspace-template/layout-rules.json`
- active pricing reference: workspace seed pricing-reference pack
  `workspace-seeds/koncept-images-pte-ltd/asset-packs/pricing-references/koncept-workspace-pricing/`

The workspace layout pack uses sanitized metadata only. It does not copy the
bundled quote-company preset containing real company header/bank details. The
quote-company profile and logo continue to come from the importable workspace
profile store, with tests using tiny sanitized logo data.

### Runtime Resolution

- Missing payload `profile_id` resolves to `koncept-workspace-template`.
- Missing payload `pricing_reference_id` resolves to
  `koncept-workspace-pricing`.
- `load_profile_pack("koncept-workspace-template")` resolves from the workspace
  seed pack path.
- `load_pricing_reference_pack("koncept-workspace-pricing")` resolves from the
  workspace seed pricing pack path.
- `/api/profiles` exposes the workspace seed pricing reference with source
  `workspace-seed` and keeps bundled repo packs separately with source
  `bundled`.
- The browser pricing-reference selector keeps non-bundled sources selectable
  and preserves the selected source in generated payloads.

### Bundled Files Demoted To Fallback/Test Fixtures

- `profiles/koncept/profile.json`
- `profiles/koncept/quotation-layout.xlsx`
- `profiles/koncept/layout-rules.json`
- `profiles/koncept/assets/koncept-header-logo.jpeg`
- `pricing-references/koncept-exhibition-quotation/reference.json`
- `pricing-references/koncept-exhibition-quotation/pricing-catalog.json`
- `pricing-references/koncept-exhibition-quotation/pricing-catalog.ai-reference.md`
- `pricing-references/koncept-exhibition-quotation/pricing-catalog-images/*`
- `fixtures/samples/kent-group/sample.json`
- `fixtures/samples/kent-group/kent-group.pdf`

`scripts/build_pricing_catalog.py` and
`scripts/validate_dynamic_pricing_reference_rules.py` remain guardrails/builders,
not deletion candidates. `_pricing-references/import-cleanup-rules.json` is an
optional ignored local import aid and is not tracked.

### Removed In Phase 3C

None. Deletion is not yet safe because the bundled Koncept profile still carries
fallback quote-company preset coverage, explicit fallback layout/pricing packs,
and existing sample/smoke coverage. Removing those files should wait until the
next deletion gate proves every fallback/test path has a smaller sanitized
replacement or is no longer needed.

### Next Deletion Gate

Before deleting the old bundled packs, a later PR must prove:

1. imported workspace quote-company profile data is mandatory or has its own
   sanitized fallback fixture
2. template/layout/rules tests no longer need `profiles/koncept`
3. pricing-reference tests no longer need
   `pricing-references/koncept-exhibition-quotation`
4. sample/PDF smoke flows no longer need `fixtures/samples/kent-group`
5. full local tests and CI pass without hidden references to removed files

Phase 3C does not add hosting, deployment, auth, Supabase, Stripe, billing,
OIDC completion, production exposure, a credit ledger, secrets, or real customer
data.

## Phase 3D Bundled Fallback Deletion Gate - 2026-06-18

Phase 3D removes the large legacy bundled logo asset and sanitizes the old
bundled quote-company fallback preset. Workspace-owned seed packs remain the
primary local/staging runtime dependencies.

### Reference Map

| Path | Phase 3D classification | Phase 3D decision |
| --- | --- | --- |
| `profiles/koncept/profile.json` | Explicit fallback/test fixture | Shrunk and sanitized. It no longer contains real header details, bank details, signatory names, or `logo_path`; the legacy preset id remains only so explicit fallback and smoke flows keep working. |
| `profiles/koncept/quotation-layout.xlsx` | Explicit fallback/test layout fixture | Kept. Generator layout tests and explicit legacy fallback coverage still depend on it. Normal workspace generation uses `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/koncept-workspace-template/quotation-layout.xlsx`. |
| `profiles/koncept/layout-rules.json` | Explicit fallback/test formatting fixture | Kept. It still covers legacy layout-rules formatting behavior and rollback tests. Normal workspace generation uses the workspace-owned layout-rules file. |
| `profiles/koncept/assets/koncept-header-logo.jpeg` | Removed legacy fallback asset | Removed. Imported workspace profile `logo_data_url` output coverage and the sanitized fallback preset data URL replace it. |
| `pricing-references/koncept-exhibition-quotation/reference.json` | Explicit fallback/test pricing metadata fixture | Kept. Active tests still cover bundled pricing-reference listing, export, import comparison, catalog matching, and fallback behavior. |
| `pricing-references/koncept-exhibition-quotation/pricing-catalog.json` | Explicit fallback/test catalog fixture | Kept. Pricing parser, matching, repair, quote basis, pricing review, and generator tests still depend on the catalog rows. |
| `pricing-references/koncept-exhibition-quotation/pricing-catalog.ai-reference.md` | Explicit fallback/test AI catalog fixture | Kept. Prompt and catalog-reference tests still cover the generated AI-facing view. |
| `pricing-references/koncept-exhibition-quotation/pricing-catalog-images/*` | Explicit fallback/test visual-reference fixture | Kept. Pricing-reference image import/export coverage still needs these visual assets. |
| `fixtures/samples/kent-group/sample.json` | Smoke/sample fixture | Kept. `/api/samples`, refresh persistence, and Playwright smoke coverage still load this sanitized sample. |
| `fixtures/samples/kent-group/kent-group.pdf` | Smoke/sample PDF fixture | Kept. PDF intake/rendering and local smoke coverage still depend on this sample PDF. |
| `_pricing-references/import-cleanup-rules.json` | Optional ignored local import cleanup aid | Runtime/import-specific spelling cleanup belongs under ignored local pricing-reference storage when needed. |
| `scripts/build_pricing_catalog.py` | Generated-reference builder | Kept. It is required for catalog import/build behavior. |
| `scripts/validate_dynamic_pricing_reference_rules.py` | Pricing hardcoding guardrail | Kept. It continues to prevent fixture-specific matching logic. |

### Removed In Phase 3D

- `profiles/koncept/assets/koncept-header-logo.jpeg`

### Shrunk Or Sanitized In Phase 3D

- `profiles/koncept/profile.json`: replaced the real bundled company fallback
  details with `Sanitized Fallback Quote Company Pte Ltd`, generic terms, a
  fixture signatory, and a tiny PNG `logo_data_url`.
- `workspace-seeds/koncept-images-pte-ltd/workspace.json`: records the old logo
  path as a removed fallback asset and points to the sanitized fallback data URL
  as replacement coverage.
- `tests/test_generate_quote.py`: generator logo-output tests now use an inline
  sanitized PNG instead of reading the deleted bundled JPEG.

### Runtime And Test Evidence

- Normal generation with omitted `profile_id` and `pricing_reference_id` is
  covered with empty legacy bundled profile/pricing roots; it resolves the
  workspace-owned template and pricing pack.
- Imported quote-company profile `logo_data_url` output coverage remains in the
  generated XLSX regression test.
- The old bundled pricing reference has been removed from tracked source.
  The Kent sample remains an explicit fixture, not a hidden runtime default.

### Next Deletion Gate

Before deleting more legacy bundled material, a later PR should migrate or
replace the remaining active fixture dependencies:

1. Move generator layout tests to a smaller sanitized layout fixture or the
   workspace-owned layout pack.
2. Replace `fixtures/samples/kent-group/kent-group.pdf` with a smaller
   deterministic sanitized PDF fixture or remove the smoke paths that require
   it.
3. Re-run the deletion gate with hidden-reference checks, focused webapp tests,
   generator tests, pricing validation, and CI before removing the kept fallback
   files.

Phase 3D does not add hosting, deployment, auth, Supabase, Stripe, billing,
OIDC completion, production exposure, a credit ledger, secrets, or real customer
data.
