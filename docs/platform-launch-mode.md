# KQAG Platform Launch Mode

This runbook covers the first KQAG-side adapter boundary for Swooshz Platform
launch handoff. It is disabled by default and does not change local/internal
KQAG mode.

## Boundary

Swooshz Platform owns login, platform sessions, users, workspaces, membership
roles, app entitlements, and app access decisions. KQAG consumes only the
platform launch context needed to create its own runtime session.

KQAG must not store provider tokens, raw provider claims, auth codes, OIDC
state, nonce, platform session cookies, raw launch tokens, or platform database
details.

## Enable Platform Launch Mode

Use placeholders for local smoke setup:

```powershell
$env:APP_MODE="deploy"
$env:AUTH_REQUIRED="true"
$env:SESSION_SECRET="<kqag-session-secret>"
$env:KQAG_PLATFORM_LAUNCH_MODE="platform"
$env:KQAG_PLATFORM_BASE_URL="https://platform.example.test"
```

`KQAG_PLATFORM_LAUNCH_MODE=disabled` is the default and keeps the existing
local/internal flow unchanged.

## Launch Consume Endpoint

KQAG accepts a platform launch token only through:

```http
POST /api/platform/launch HTTP/1.1
X-App-Launch-Token: <one-time-platform-launch-token>
```

The adapter then calls:

```http
POST <platform-base-url>/api/platform/apps/launch/consume?appKey=kqag
X-App-Launch-Token: <one-time-platform-launch-token>
```

The raw token must not be placed in query parameters, browser storage, logs,
files, screenshots, docs, or telemetry. After consume, KQAG stores only the safe
platform context returned by the consume response in its signed KQAG session.

## Accepted Consume Context

KQAG stores only these fields:

- `outcome`
- `user.userId`
- `user.email`
- `user.displayName`
- `user.status`
- `workspace.workspaceId`
- `workspace.workspaceSlug`
- `workspace.workspaceName`
- `app.appKey`
- `app.appName`
- `membershipRole`
- `launchTokenExpiresAt`

KQAG rejects missing tokens, consume failures, non-`consumed` outcomes, wrong
app keys, missing platform user IDs, missing workspace IDs, and stale expiry
values.

## Deferred Work

This adapter does not add cloud or database-backed KQAG storage. Platform-scoped
storage for profiles, pricing references, and past sessions belongs in the next
PR after the launch context boundary has been proven.
