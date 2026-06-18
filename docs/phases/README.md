# Quote Generator Phase Index

This folder tracks SKR-style execution phases for the Swooshz quote generator.
Keep `docs/mvp-implementation-plan.md` as the canonical architecture source of
truth for platform auth, company data, credits, billing, deployment, and product
boundaries. Use these phase docs to sequence PR-sized work against that plan.

## Phase Order

- [Phase 0 - Roadmap split and PR check system](phase-0-roadmap-split-and-pr-checks.md)
- [Phase 1 - Protected internal usable copy](phase-1-protected-internal-usable-copy.md)
- [Phase 2 - Koncept Images Pte Ltd first internal workspace seed](phase-2-koncept-images-workspace-seed.md)
- [Phase 3 - Safe export/import verification and repo cleanup](phase-3-export-import-verification-and-cleanup.md)
- [Phase 4 - Hostinger/Coolify protected staging deployment](phase-4-hostinger-coolify-protected-staging.md)
- [Phase 5 - Internal team testing and feedback loop](phase-5-internal-team-testing-feedback.md)
- [Phase 6 - Real platform auth, company, user, role, credit ledger, and billing foundation](phase-6-platform-foundation.md)
- [Phase 7 - Production MVP hardening](phase-7-production-mvp-hardening.md)

## PR Rule

Every future PR should identify its phase and subphase in
`docs/pr-checks/quote-generator-pr-checklist.md` format. If a PR crosses phase
boundaries, explain why and keep the first deployable behavior slice small.

## PR #28 Follow-Up Rule

PR #28 added quote-company profile preset Save, Delete, Import, and Export
controls, separated saved profiles from read-only templates, improved pricing
reference management controls, and kept saved profile defaults sanitized before
persistence.

Those profile import/export controls are now the preferred bridge for creating
the first Koncept Images Pte Ltd internal workspace seed. Future implementation
should not add random hardcoded company data when the export/import path can be
used instead. Exported company profile and pricing artifacts should be treated
as portable seed inputs that can later migrate into the platform company model.

Repo-bundled templates, pricing references, profile fixtures, and generated
reference builders should be removed only after export/import verification and
test migration prove the replacement path is safe.
