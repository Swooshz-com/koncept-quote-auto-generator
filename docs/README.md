# KQAG Documentation Index

This index describes the active documentation set for the internal KQAG/SAQG
quote-generator module. KQAG owns the quote-specific local workflow. Future
platform concerns belong in `Swooshz-com/swooshz-platform`.

## Current Docs

- `docs/kqag-current-status.md`: current RC verdict, module boundary, runtime
  data model, private asset rules, owner smoke checklist, audit findings, and
  docs cleanup summary.
- `docs/testing-plan.md`: validation expectations for product, frontend,
  import/export, security, CI, and smoke-test changes.
- `docs/internal-uat.md`: internal Koncept/Swooshz UAT checklist, smoke
  commands, known limits, bug-report format, and private-data guardrails.
- `docs/pricing-catalog-import.md`: current pricing-reference import behavior,
  AI normalization/enrichment contracts, save behavior, ordering, and deferred
  import items.
- `docs/ai-basis-chat-test-playbook.md`: AI basis chat test scope, mocked checks,
  live smoke guidance, and prompt/response expectations.
- `docs/privacy-pdpa-gdpr-baseline.md`: privacy and legal engineering baseline.
  Treat production and account-related entries there as future launch/platform
  blockers, not internal RC implementation work.
- `docs/current-cicd-status.md`: active GitHub Actions workflow, CI checks, and
  maintenance rule for CI/CD changes.
- `docs/pr-checks/quote-generator-pr-checklist.md`: PR review checklist and
  KQAG module boundary reminder.
- `docs/agent-playbooks/`: portable AI-agent playbooks referenced by `AGENTS.md`.

## Historical Or Archive Docs

No historical/archive docs remain in this cleanup pass. Completed phase plans and
old handoff notes were consolidated instead of archived.

## Removed Or Consolidated Docs

- `docs/internal-team-test-handoff.md` was consolidated into
  `docs/kqag-current-status.md` and removed.

No unique current RC requirements were deleted without being summarized in the
current status/handoff doc.

## Future Platform Ownership

The following topics should not be implemented in this repo as part of KQAG RC:

- login, accounts, users, roles, company membership, and app access
- Stripe, billing, credits, subscriptions, ledgers, and entitlement
- Supabase or other hosted database design
- DB-backed quote history and dashboards
- Hostinger, Coolify, Docker deployment, DNS, public hosting, or production
  infrastructure
- Swooshz platform shell, navigation, app registry, app whitelist, and cross-app
  architecture
- SEOzilla integration or other platform-level app integrations

## Runtime And Private Asset Reminder

Internal KQAG testing should use clean local/runtime storage:

```powershell
QUOTE_DATA_ROOT=<local-runtime-data-root>
```

Private profile and pricing files stay outside the repository:

- `<private-profile-json-outside-repo>`
- `<private-pricing-xlsx-outside-repo>`

Do not commit runtime stores, private uploads, generated workbooks, generated
PDFs, real logos, workbook rows/cells, bank/payment data, customer data, or
private local paths.
