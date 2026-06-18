# Quote Generator PR Checklist

Copy this checklist into every future quote-generator PR and complete each
section before requesting review.

## Required PR Fields

- Phase and subphase:
- Objective:
- Out of scope:
- User-facing change:
- UI files touched (`webapp/static/index.html`, `webapp/static/styles.css`,
  visible layout/DOM/workflow portions of `webapp/static/app.js`, Playwright
  screenshot/smoke expectations): Yes/No
- Explicit current-turn UI approval if UI files were touched:
- Data, storage, or auth impact:
- Live/external service or deployment action involved: Yes/No
- Secrets or env values touched: Yes/No
- Bundled files, templates, pricing references, profiles, fixtures, or generated
  outputs touched: Yes/No
- Import/export behavior touched: Yes/No
- Company/workspace seed data touched: Yes/No
- Validation commands and results:
- Manual smoke test evidence when relevant:
- Rollback notes when relevant:
- Remaining risks:

## Required Review Notes

- SAQG/KQAG solution UI is complete and final. Do not change visible UI,
  layout, DOM, CSS, workflow placement, cards, tabs, buttons, modals, spacing,
  component hierarchy, or visual status components unless the user explicitly
  approves UI work in the current turn. Text-only wording changes and
  backend/data mapping are allowed when scoped.
- If UI files are touched without explicit current-turn UI approval, the PR
  must fail/reject itself unless the change is text-only wording or invisible
  data serialization.
- If live/external service, deployment, Docker, Hostinger, Coolify, Supabase,
  Stripe, OIDC, DNS, firewall, env var, secret, credential, or production action
  is involved, state the explicit approval and target operation.
- If secrets or env values are touched, confirm no secret values are committed,
  logged, screenshotted, or exposed to frontend/static code.
- If bundled files, templates, pricing references, profiles, fixtures, or
  generated outputs are touched, state whether they are source files,
  generated files, seed artifacts, or test fixtures.
- If import/export behavior is touched, state the export source, import target,
  validation path, and whether imported data remains migratable to the platform
  company model.
- If company/workspace seed data is touched, state whether it is real internal
  company data, fixture data, or generated test data.
- If deleting repo-bundled assets, confirm export is verified, import is
  verified, tests are migrated, and no active guardrail or canonical doc is
  removed.

## PR #28 Follow-Up Requirement

For Koncept Images Pte Ltd seed work, prefer the profile import/export and
pricing-reference export/import paths added by PR #28. Do not add arbitrary
hardcoded company data when the exported artifact path can supply the seed.
Treat exported profile and pricing artifacts as portable seed inputs for the
later platform company model.
