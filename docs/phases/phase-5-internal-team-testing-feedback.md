# Phase 5 - Internal Team Testing And Feedback Loop

## Goal

Use the protected internal copy in practical team workflows while the full
platform is still being built.

## Scope

- Define feedback capture for bugs, confusing flows, missing pricing rows,
  import/export issues, and quote-output formatting issues.
- Run repeatable smoke tests after meaningful app, seed, or deployment changes.
- Triage feedback into fix-now, document, defer, and platform-phase buckets.
- Keep safe-use notes visible for internal testers.

## Safe-Use Notes

- Use approved internal test jobs unless real customer data is explicitly
  approved for the current task.
- Do not treat local/staging storage as durable production storage.
- Do not share generated files externally unless the business owner approves the
  specific quote/customer context.
- Report pricing uncertainty through the pricing review flow or the PR's
  remaining-risk notes.

## Smoke Test Areas

- Protected access path.
- Image/PDF upload and analysis readiness.
- Quote-company profile import/export.
- Pricing-reference export/import/edit where relevant.
- Quote basis review and pricing review.
- Excel download and basic workbook formatting.
- Logs under `_logs/` only.

## Exit Criteria

- Internal testers have a lightweight feedback path.
- Bugs and improvement requests are triaged against phase scope.
- A small smoke checklist is reused across internal testing PRs.
