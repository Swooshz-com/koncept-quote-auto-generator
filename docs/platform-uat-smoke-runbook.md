# KQAG Platform UAT Smoke Runbook

This runbook proves the first internal Swooshz Platform to KQAG handoff path:
Platform owns login, session, workspace, membership, role, entitlement, and app
access decisions; KQAG consumes the Platform launch context, creates its own
runtime session, and stores quote sessions and generated artifacts under the
Platform workspace context.

Use placeholders only in notes and reports. Do not paste real launch tokens,
provider tokens, auth codes, OIDC state, nonce values, provider responses,
callback URLs with query parameters, database URLs, staff emails, production
domains, private profile paths, or generated customer quotes into chat, docs,
issues, logs, screenshots, or PR text.

## Scope

In scope:

- Platform Google OIDC sign-in and Platform session creation.
- Platform KQAG app access and launch intent creation.
- KQAG launch consume through `POST /api/platform/launch`.
- Raw launch token forwarded from KQAG to Platform only in the
  `X-App-Launch-Token` header.
- KQAG signed runtime session with safe Platform context only.
- Platform-scoped KQAG database rows for quote sessions.
- Platform-scoped generated XLSX artifact storage and download through the
  quote-session route used by the past-session dashboard.

Out of scope:

- Production deployment, DNS, TLS, or reverse proxy setup.
- Public signup, billing, invitations, member management, or admin dashboard.
- KQAG-owned accounts, login, membership, billing, or app registry.
- Object storage and multi-instance durable job execution.
- Private Koncept profile or pricing files committed to the repo.

## Automated Local Coverage

The unit test
`test_platform_uat_smoke_launch_generate_list_and_download_database_artifact`
uses a fake Platform consume response, a synthetic generator subprocess, and a
temporary SQLite database. It proves that KQAG can:

- Accept a launch token only through `X-App-Launch-Token`.
- Call Platform consume with the raw token only in the header.
- Create a KQAG session from safe Platform context.
- Generate a synthetic XLSX through the normal `/api/generate` HTTP path.
- Persist quote session and artifact rows under the Platform workspace ID.
- List the generated session through `/api/quote-sessions`.
- Download the stored XLSX through the quote-session download route.
- Avoid storing or returning the raw launch token.

Run the focused smoke locally with:

```powershell
python -m unittest tests.test_webapp.WebappServerTest.test_platform_uat_smoke_launch_generate_list_and_download_database_artifact
```

## Local Environment

KQAG Platform launch mode is explicit and disabled by default:

```powershell
$env:APP_MODE="deploy"
$env:AUTH_REQUIRED="true"
$env:SESSION_SECRET="<kqag-session-secret>"
$env:KQAG_PLATFORM_LAUNCH_MODE="platform"
$env:KQAG_PLATFORM_BASE_URL="<platform-base-url>"
$env:KQAG_STORAGE_MODE="database"
$env:KQAG_ARTIFACT_STORAGE_MODE="database"
$env:KQAG_DATABASE_URL="<kqag-local-database-url>"
```

Apply the reviewed KQAG storage migrations only against a disposable local
database:

```powershell
@'
from webapp import server
server.apply_kqag_storage_migrations("<kqag-local-database-url>")
'@ | python -
```

Start KQAG locally after the environment is set:

```powershell
python -m webapp.server
```

## Manual Platform Smoke

1. In the Platform repo, complete the reviewed local pre-smoke commands and
   migrations against a disposable local database.
2. Start Platform with `npm run platform:start`.
3. Open `<platform-base-url>/` in the browser.
4. Complete Google sign-in.
5. Confirm the callback lands on `/app`.
6. Seed the logged-in provider-backed user for KQAG access using the Platform
   seed command with the email used for sign-in. Do not paste the real email
   into chat or PR text.
7. Refresh `/app` and confirm session context, workspace, and KQAG app access
   appear.
8. Create a KQAG launch intent in the Platform shell.
9. Keep the raw launch token in the temporary handoff area only long enough to
   submit the KQAG launch request. Do not put it in a URL or browser storage.
10. In a local terminal, submit the token to KQAG as a header-only POST. This
    creates a KQAG session in the PowerShell web session, not in the browser:

```powershell
$launchToken = "<one-time-platform-launch-token>"
$launchResponse = Invoke-WebRequest `
  -Method Post `
  -Uri "<kqag-base-url>/api/platform/launch" `
  -Headers @{ "X-App-Launch-Token" = $launchToken } `
  -SessionVariable kqagSession
Remove-Variable launchToken
```

11. Confirm the terminal-scoped KQAG session context loads:

```powershell
Invoke-WebRequest `
  -Method Get `
  -Uri "<kqag-base-url>/api/session" `
  -WebSession $kqagSession
```

12. For generated quote storage, session listing, and XLSX download coverage
    without private files or live services, run the automated local smoke test
    above. Use safe test data only for any additional manual generation.
13. Confirm no raw launch token, provider token, auth code, OIDC state, nonce,
    provider payload, callback query, database URL, staff email, production
    domain, private local path, or generated customer quote appears in logs,
    browser storage, screenshots, docs, or PR text.

## Browser Launch Limitation

The current Platform internal shell intentionally displays the one-time launch
handoff payload and does not call KQAG. KQAG intentionally accepts the raw
launch token only in the `X-App-Launch-Token` header. A direct cross-origin
browser navigation cannot attach that header, and putting the token in a query
string or browser storage is forbidden.

The terminal handoff above intentionally does not claim a normal browser has a
KQAG session cookie. The current strongest safe local coverage is the automated
KQAG smoke test, plus the manual Platform Google OIDC and launch-intent check.
A fully clickable browser launch should be added later through a reviewed
same-origin/reverse-proxy or similarly safe handoff design that preserves the
header-only consume rule and does not expose the raw token in URLs, storage,
logs, or screenshots.

## Pass Criteria

- Platform session context loads after Google sign-in.
- Platform workspace appears.
- KQAG app access appears.
- Platform launch intent can be created.
- KQAG launch consume succeeds.
- KQAG terminal-scoped session context loads with Platform workspace context.
- Automated local smoke proves a generated quote session is stored under the
  Platform workspace ID.
- Automated local smoke proves a generated XLSX artifact downloads from the
  quote-session route.
- Raw launch token is not stored after consume and is not present in reports,
  screenshots, logs, URLs, browser storage, database rows, or KQAG responses.
