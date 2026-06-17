# Phase 7 - Production MVP Hardening

## Goal

Prepare the MVP for production use after platform auth, company isolation,
credits, billing, and protected staging have been proven.

## Scope

- Durable storage for uploads and generated files.
- Observability for app health, job failures, AI usage, and security-relevant
  events.
- Security hardening for auth, session handling, upload limits, rate limits,
  secrets, and company isolation.
- Backup and restore procedures for persistent data.
- Deployment gates, rollback paths, and release evidence.
- Production smoke tests that cover upload-to-XLSX and company/user/role/credit
  boundaries.

## Production Gates

- Auth complete and tested.
- HTTPS complete.
- Company isolation and RLS tested.
- User invite/deactivation and role changes tested.
- AI usage caps and credit ledger active.
- Billing webhook and customer portal tested in the intended mode.
- Generated files stored durably.
- Logs avoid secrets and customer file contents.
- Backup/restore and rollback procedures documented.
- Health checks and smoke tests pass.

## Exit Criteria

- Production readiness is documented with evidence.
- Deployment can be rolled back without losing durable customer data.
- The MVP no longer depends on local-only storage, temporary company bridges, or
  staging-only protection as production controls.
