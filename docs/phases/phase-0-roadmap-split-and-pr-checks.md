# Phase 0 - Roadmap Split And PR Check System

## Goal

Convert the broad MVP plan into execution phases and establish a repeatable PR
checklist for quote-generator changes.

## Scope

- Keep `docs/mvp-implementation-plan.md` as the canonical architecture source of
  truth.
- Add SKR-style phase docs that sequence work without changing app behavior.
- Add a PR checklist that forces future PRs to state phase, scope, data impact,
  auth impact, import/export impact, validation, rollback notes, and risks.
- Make PR #28 follow-up explicit: the profile import/export controls are the
  preferred bridge for the first Koncept Images Pte Ltd internal workspace seed.

## Out Of Scope

- No app behavior change.
- No Docker, Hostinger, Coolify, Supabase, Stripe, OIDC, auth, billing, or credit
  ledger implementation.
- No deletion of bundled files, templates, profiles, pricing references,
  fixtures, validation scripts, or generated-reference builders.
- No secrets or real customer data.

## Exit Criteria

- `docs/phases/README.md` links every phase.
- Every phase has a short, reviewable phase doc.
- `docs/pr-checks/quote-generator-pr-checklist.md` exists and can be copied into
  future PR descriptions.
- Validation confirms the docs-only diff has no whitespace errors.
