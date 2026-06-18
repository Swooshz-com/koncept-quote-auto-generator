# Phase 2 - Koncept Images Pte Ltd First Internal Workspace Seed

## Goal

Create the first real internal company/workspace identity for Koncept Images Pte
Ltd using the import/export foundation from PR #28.

## Scope

- Treat Koncept Images Pte Ltd as the first real internal company/workspace
  identity, not disposable demo data.
- Use the current app's exported quote-company profile, pricing reference, and
  template artifacts as portable seed inputs.
- Allow an early local or staging company-data bridge only if it is designed to
  migrate into the later Supabase/platform company model.
- Keep bundled profile and pricing-reference files available until export/import
  verification and test migration are complete.

## Seed Source Rule

The newly merged profile import/export controls are the preferred bridge for
creating the first Koncept Images Pte Ltd internal workspace seed. Future work
should not add random hardcoded company data when a profile/pricing export can
be produced by the app and imported into the bridge.

Exported company profile and pricing artifacts should be treated as portable
seed inputs. A file such as
`docs/koncept-images-pte-ltd.quote-company-profile.json` is an example of the
intended exported profile artifact from the PR #28 notes, but this phase must
not require that exact untracked local file to exist.

## Migration Requirement

Any temporary local/staging seed storage must preserve enough identity and data
shape to migrate later into platform-owned records such as:

- company identity
- quote-company profile presets
- pricing reference packs and versions
- selected profile/template relationships
- users, memberships, roles, entitlements, and credit ledger links once Phase 6
  begins

## Out Of Scope

- No Supabase schema implementation.
- No production auth, Stripe billing, real credit ledger, or public onboarding.
- No hardcoded company branches in source code.
- No deletion of bundled profile/pricing/template files.

## Exit Criteria

- The seed path is documented and driven by export/import artifacts.
- The Koncept Images Pte Ltd identity is represented as internal company/workspace
  seed data, not as sample-only fixture logic.
- The PR states how the seed will migrate into the platform company model later.

## Phase 2A Local/Staging Bridge

The first internal workspace seed is represented by
`workspace-seeds/koncept-images-pte-ltd/workspace.json`:

- display name: `Koncept Images Pte Ltd`
- stable company/workspace id: `koncept-images-pte-ltd`
- local profile preset storage target:
  `QUOTE_DATA_ROOT/koncept-images-pte-ltd/profiles.json`
- local pricing-reference storage target:
  `QUOTE_DATA_ROOT/koncept-images-pte-ltd/pricing-references.json`
- active Phase 3C workspace layout/template pack:
  `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/koncept-workspace-template/`
- active Phase 3C workspace pricing-reference pack:
  `workspace-seeds/koncept-images-pte-ltd/asset-packs/pricing-references/koncept-workspace-pricing/`

This is a bridge toward the later platform company model only. It does not add
Supabase, production auth, billing, a credit ledger, deployment, or public
onboarding. The seed manifest preserves the company id, display name, selected
quote-company profile store, workspace-owned layout/template and pricing pack
ids, storage collections, created/updated metadata, and migration notes so
imported profile presets and workspace asset packs can later become
company-scoped platform records.

## Maintainer Export/Import Notes

1. In the local app, open the Quote Company panel and load the current Koncept
   Exhibition/Koncept Images profile preset.
2. Use the quote-company profile Export control to save a
   `*.quote-company-profile.json` artifact.
3. Use the quote-company profile Import control to select that artifact. Confirm
   that the company header text and logo preview are restored.
4. Save the imported profile. The app stores it under the active
   `koncept-images-pte-ltd` local workspace path instead of the old placeholder
   `default` company folder.

The bridge does not require any untracked local file such as
`docs/koncept-images-pte-ltd.quote-company-profile.json` to exist. That file can
be used as a local maintainer export artifact for smoke testing, but it should
not be committed if it contains real internal/customer data.
