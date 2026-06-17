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
