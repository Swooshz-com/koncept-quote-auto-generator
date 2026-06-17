# Phase 4 - Hostinger/Coolify Protected Staging Deployment

## Goal

Plan and implement a protected staging deployment on Hostinger VPS + Coolify
after the internal seed path is safe.

## Scope

- Add Docker/Coolify setup only in this later phase.
- Keep staging protected and clearly separate from production readiness.
- Plan persistent paths for generated files, temporary files, logs, and future
  durable storage.
- Add smoke tests and deployment evidence requirements before exposing a staging
  URL to testers.

## Non-Negotiables

- Staging-only protection is required before any external access.
- No public database, cache, admin port, or unprotected quote-generator endpoint.
- No secrets in repo files, static frontend code, logs, screenshots, or PR text.
- No production deployment claim until Phase 7 gates are complete.

## Planned Evidence

- Coolify service names and runtime boundaries.
- Persistent mount plan for quote output, temporary files, and logs.
- Health-check URL and smoke-test steps.
- Rollback path.
- Confirmation that no deployment job was added to CI unless this phase
  explicitly updates and documents that boundary.

## Exit Criteria

- Protected staging can be deployed, smoke tested, rolled back, and documented.
- CI/CD status documentation matches any workflow or deployment changes.
