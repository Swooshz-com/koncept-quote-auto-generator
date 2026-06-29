# KQAG Platform-Scoped Storage Mode

This runbook covers the first KQAG-owned app-data storage boundary for platform
team mode. It is disabled by default. Existing local/internal mode continues to
use local runtime storage under `QUOTE_DATA_ROOT` and related local roots.

## Boundary

Swooshz Platform owns users, login, platform sessions, workspaces, membership
roles, app access, invites, and billing. KQAG stores only quote-generator app
data scoped to the platform workspace ID present in the signed KQAG platform
session.

KQAG storage must not store raw platform launch tokens, provider tokens, raw
provider claims, auth codes, OIDC state, nonce, platform cookies, platform
session secrets, database passwords in logs, private local paths, or private
profile/pricing files in Git.

## Modes

Local mode is the default:

```powershell
$env:KQAG_STORAGE_MODE="local"
```

Database mode requires a valid platform-launched KQAG session and a configured
KQAG database URL:

```powershell
$env:KQAG_PLATFORM_LAUNCH_MODE="platform"
$env:KQAG_PLATFORM_BASE_URL="https://platform.example.test"
$env:KQAG_STORAGE_MODE="database"
$env:KQAG_DATABASE_URL="sqlite:///C:/path/to/local/kqag-storage.sqlite3"
```

The first implementation supports the reviewed SQLite migration path. Unsupported
database URL schemes fail closed with a generic app-facing storage error and
privacy-safe logs.

## Migration

Review `migrations/001_platform_scoped_storage.sql`, then apply it explicitly:

```powershell
$env:KQAG_DATABASE_URL="sqlite:///C:/path/to/local/kqag-storage.sqlite3"
python scripts/migrate_kqag_storage.py
```

The app does not auto-run migrations on startup.

## Workspace Scope

Database rows are keyed by the platform workspace ID from the KQAG platform
session. Profiles, pricing references, and quote sessions saved by workspace A
must not list, read, export, or delete from workspace B.

## Included App Data

The boundary covers:

- quote-company profile list, save, delete, and export payload resolution
- pricing-reference list, detail, save, delete, and export payload resolution
- quote-session list, read, save, delete, and download metadata resolution

Local profile/pricing/session behavior remains the default and continues to use
existing runtime storage.

## Out Of Scope

This does not add a platform admin dashboard, invites or member management,
KQAG-owned login/auth, fake login, billing, Stripe, deployment, DNS/TLS, public
signup, object storage, private profile/pricing files, or generated customer
quotes in Git.
