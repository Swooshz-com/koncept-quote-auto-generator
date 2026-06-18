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
- `pricing-references/import-cleanup-rules.json`
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
