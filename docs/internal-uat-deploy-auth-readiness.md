# Internal UAT Deploy/Auth Readiness

## Purpose

This document explains the existing single-instance gated internal UAT path for
KQAG/SAQG deploy/auth scaffolding. It is a repo-specific readiness guide for
Koncept/Swooshz internal testing only.

This document does not approve a production launch, public SaaS access,
customer portal access, ecommerce, billing, DB-backed multi-user mode, or full
Swooshz platform integration. It also does not move platform-owned account,
membership, legal, billing, or app-whitelist work into KQAG.

## Current Repo-Supported Knobs

The deploy/auth surface is already represented in `.env.example` and
`webapp/server.py`.

- `APP_MODE=local`: localhost-first desktop/dev mode. The server defaults to
  loopback binding, local host allowlisting, and `AUTH_REQUIRED=false` unless
  explicitly overridden.
- `APP_MODE=deploy`: gated deploy mode. The server defaults to a deploy bind
  host and `AUTH_REQUIRED=true`.
- `AUTH_REQUIRED`: explicit auth gate toggle. In deploy mode, leaving this
  unset still defaults to auth required.
- `SESSION_SECRET`: required for signed session and OIDC state cookies when
  deploy auth is enabled. Never print or commit the value.
- `OIDC_ISSUER_URL`, `OIDC_CLIENT_ID`, `OIDC_CLIENT_SECRET`,
  `OIDC_REDIRECT_URI`, and `OIDC_LOGOUT_URL`: OIDC settings used by the
  deploy auth scaffold. `OIDC_LOGOUT_URL` is optional; the other OIDC fields
  plus `SESSION_SECRET` are required for a complete auth boundary.
- `QUOTE_DATA_ROOT`: runtime company/profile/pricing/session data root.
- `QUOTE_OUTPUT_ROOT`: generated quote output root.
- `QUOTE_TMP_ROOT`: temporary job/work root.
- `QUOTE_LOG_ROOT`: runtime log root.
- `USER_TYPE`: local role simulation for desktop/internal testing. Deploy mode
  must rely on authenticated session claims rather than local role simulation.

The server also changes cookie behavior in deploy mode: signed session and OIDC
state cookies are emitted with `Secure`, `HttpOnly`, and `SameSite=Lax`.

## Current Implementation Notes

- Deploy mode is intended to require authentication by default.
- Deploy mode refuses to start when auth is required and the auth boundary is
  incomplete.
- OIDC configuration completeness is checked before serving authenticated
  deploy traffic.
- `/login` redirects to the configured OIDC issuer when the auth boundary is
  complete.
- The OIDC callback scaffold currently returns `501 not_implemented` until
  token exchange and claims validation are wired. Treat this as a readiness
  gate before any broader deploy UAT with authenticated testers.
- In-memory jobs are acceptable for local mode and a first single-instance UAT
  deploy only. Multi-instance deployment requires durable job, upload, download,
  log, pricing-reference, and quote-session storage partitioned by authenticated
  user/account.

## Recommended Internal UAT Shape

Use this shape only for gated internal UAT after the auth callback readiness
gate is satisfied:

- Single UAT host.
- Single app instance.
- `APP_MODE=deploy`.
- `AUTH_REQUIRED=true`.
- Complete OIDC configuration.
- Persistent runtime data root outside the repository.
- Approved tester access only.
- No public/customer access.
- No multi-instance scaling.
- No committed runtime data.
- No real secrets in repository files.

For quote workflow coverage that does not depend on deploy auth, keep using the
local internal UAT checklist in `docs/internal-uat.md`.

## Explicitly Not Ready

The current KQAG repo is not ready for:

- Production launch.
- Public SaaS.
- Customer portal access.
- Multi-instance deployment.
- DB-backed session/history.
- Durable per-user storage partitioning.
- Billing, credits, checkout, orders, or ecommerce.
- User, account, team, membership, or customer-management models.
- Platform-owned app whitelist/account membership.
- Full OIDC token exchange and claims validation for public use.
- Legal/customer-facing production launch without privacy, terms, retention,
  sub-processor, and counsel review.

## UAT Deploy Smoke Checklist

Use this checklist for the gated internal UAT deploy path. Do not print secrets
while checking env values.

- [ ] Confirm required env names are present without printing values:
      `APP_MODE`, `AUTH_REQUIRED`, `SESSION_SECRET`, `OIDC_ISSUER_URL`,
      `OIDC_CLIENT_ID`, `OIDC_CLIENT_SECRET`, `OIDC_REDIRECT_URI`, and the
      runtime root envs being used.
- [ ] Confirm `APP_MODE=deploy`.
- [ ] Confirm `AUTH_REQUIRED=true`.
- [ ] Confirm runtime roots are outside the repository.
- [ ] Confirm the app refuses unsafe or incomplete deploy-auth configuration.
- [ ] Confirm the health endpoint responds.
- [ ] Confirm unauthenticated users are blocked or redirected.
- [ ] Confirm OIDC login redirects to the configured issuer.
- [ ] Confirm the OIDC callback readiness gate is satisfied before treating the
      deploy UAT as authenticated end-to-end. In the current scaffold, a
      `501 not_implemented` callback response means token exchange and claims
      validation still need a separate implementation PR.
- [ ] Confirm an authenticated approved tester can reach the dashboard.
- [ ] Confirm New Quote works.
- [ ] Confirm profile/pricing import works with authorised private local files.
- [ ] Confirm quote generation works.
- [ ] Confirm XLSX/PDF export works.
- [ ] Confirm dashboard modify/download/delete works.
- [ ] Confirm direct runtime/output files are not publicly browsable.
- [ ] Confirm `git status --short` has no private runtime/output files.

## Private Data And Secret Rules

Do not commit, paste into GitHub, print in logs, or include in screenshots:

- `.env` secrets.
- OIDC client secret.
- Session secret.
- Tunnel/provider tokens.
- Real profile JSON.
- Files containing `logo_data_url`.
- Embedded Base64 logos.
- Real pricing files.
- Generated quote exports.
- Runtime session folders.
- Customer, company, or bank data.
- Private filesystem paths.

Use placeholders in docs, issues, PRs, and bug reports, and redact screenshots
unless they use clearly synthetic/test-only data.

## Validation Evidence To Keep With UAT Notes

For each internal deploy-auth UAT run, record only non-secret evidence:

- Date/time and app version or commit.
- `APP_MODE` value.
- Whether `AUTH_REQUIRED` is enabled.
- Whether required OIDC env names are set, without values.
- Runtime root category, such as `outside repo`, without private paths.
- Health endpoint result.
- Auth redirect/block result.
- OIDC callback result.
- Quote workflow smoke result if authenticated dashboard access is available.
- Confirmation that no private runtime/output files appear in `git status`.
