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
