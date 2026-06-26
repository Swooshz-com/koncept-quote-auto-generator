# KQAG Internal UAT Coolify Deploy Adapter

## Purpose

This adapter is for running KQAG/SAQG as a bounded single-instance internal UAT
app on an already-prepared Coolify host.

Primary workspace/company: Koncept Images Pte Ltd.
Workspace slug: `koncept-images-pte-ltd`.

This document is not a generic Hostinger, VPS, Coolify, SSH, firewall, DNS, or
server-maintenance guide. Toolkit-owned infrastructure guidance remains outside
this repo. KQAG owns only the app-specific internal UAT requirements listed
here.

This adapter does not approve production launch, public SaaS, customer portal
access, ecommerce, billing, DB-backed multi-user mode, multi-instance scaling,
or full Swooshz platform integration.

## App Shape

Use this shape only for internal UAT:

- One prepared Coolify host.
- One KQAG application.
- One running app instance.
- Python buildpack or equivalent repo build using `requirements.txt`.
- Start command: `python webapp/server.py`.
- `APP_MODE=deploy`.
- `AUTH_REQUIRED=true`.
- Complete OIDC settings, including explicit `OIDC_AUTHORIZE_URL`,
  `OIDC_TOKEN_URL`, and `OIDC_USERINFO_URL`.
- Approved tester allowlist through `AUTH_ALLOWED_EMAILS` and/or
  `AUTH_ALLOWED_DOMAINS`.
- Persistent runtime storage mounted outside the repository path.
- No public/customer access.
- No real secrets, runtime files, profile exports, pricing files, or generated
  quote exports committed to git.

The server reads `PORT` and binds to `0.0.0.0` automatically when
`APP_MODE=deploy`, so a separate app-specific Dockerfile is not required for the
current internal UAT path.

## Coolify App Settings

Use the already-prepared Coolify host and create a single KQAG app from this
repository.

- Build/install command: install Python dependencies from `requirements.txt`
  using the platform's Python buildpack defaults.
- Start command: `python webapp/server.py`.
- Port: `8765`, or the value supplied by `PORT`.
- Healthcheck path: `/api/health`.
- Healthcheck expected result: HTTP `200` with JSON status `ok`.
- Instance count: `1`.
- Scaling: disabled for this UAT adapter.

Do not add a database, queue, object store, external session store, Supabase
project, platform shell, app registry, billing service, or customer portal for
this UAT adapter.

## Environment Template

Use `deploy/internal-uat/coolify/kqag.uat.env.example` as the app-specific
placeholder checklist. Copy values into Coolify secrets/environment management;
do not commit a populated `.env`.

Before a VPS is available, verify the placeholder template locally:

```powershell
python scripts\verify_internal_uat_deploy_template.py
```

Required UAT checks before starting the app:

```powershell
python webapp\server.py --check-deploy-uat-env
```

The preflight reports only check names/messages and must not print secret
values.

Approved internal tester login expectations and the full pre-VPS dry-run scope
are documented in `docs/internal-uat-login-and-pre-vps-dry-run.md`.

## Persistent Runtime Storage

Map persistent storage for the four runtime roots in the env template:

- `QUOTE_DATA_ROOT=/var/lib/kqag/data`
- `QUOTE_OUTPUT_ROOT=/var/lib/kqag/output`
- `QUOTE_TMP_ROOT=/var/lib/kqag/tmp`
- `QUOTE_LOG_ROOT=/var/log/kqag`

See `deploy/internal-uat/coolify/volume-map.example.md` for the repo-specific
volume map. These paths are runtime data, not source files. They must not be
committed, exposed as static files, or browsable through the public app.

## Smoke Checklist

Run this after Coolify deploys the single UAT app. Do not print secrets while
checking env values.

- [ ] App deploy/build succeeds.
- [ ] App starts.
- [ ] `/api/health` works.
- [ ] Unauthenticated browser request redirects to login.
- [ ] Unauthenticated API request blocks.
- [ ] OIDC login redirects to exact configured `OIDC_AUTHORIZE_URL`.
- [ ] Approved tester reaches dashboard.
- [ ] Unapproved tester is blocked generically.
- [ ] New Quote works.
- [ ] Profile import works with an authorised private local file.
- [ ] Pricing reference import works with an authorised private local file.
- [ ] Quote generation works.
- [ ] XLSX/PDF export works.
- [ ] Dashboard modify/download/delete works.
- [ ] Runtime/output/tmp/log files are not publicly browsable.
- [ ] Logs do not reveal secrets, auth codes, tokens, provider responses,
      private data, or raw customer/company/bank data.
- [ ] `git status --short` is clean of runtime/private files.

## Boundary

Toolkit or infrastructure runbooks own:

- Hostinger/VPS setup.
- Coolify installation and upgrades.
- SSH access and owner approvals.
- Firewall, DNS, TLS, and network exposure.
- Server maintenance evidence and infrastructure workflow.

KQAG owns:

- App start command.
- App healthcheck path.
- App deploy-mode env shape.
- App runtime root names and volume expectations.
- Internal UAT auth/smoke checklist.
- Private-data and secret guardrails for KQAG quote workflow files.

Future production platform/accounts/billing/DB/customer portal work remains out
of scope for this repo.
