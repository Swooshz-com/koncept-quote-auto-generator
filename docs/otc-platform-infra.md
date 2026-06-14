# OTC AI Platform Infrastructure

## Target Platform

Use **Hostinger VPS + Coolify** for the first hosted deployment.

Keep Supabase or S3-compatible managed storage available for durable customer files. Keep managed Supabase/Postgres available for customer data until self-hosted backup and restore is tested.

## Server

Minimum staging server:

```text
2 vCPU
8 GB RAM
80 GB NVMe
Ubuntu LTS
Docker
Coolify
```

Preferred early production server:

```text
4 vCPU
8-16 GB RAM
150 GB+ NVMe
Ubuntu LTS
Docker
Coolify
```

## Coolify Apps

### `otc-frontend`

- Hosts the final user-facing frontend.
- Serves quote workflow pages.
- Handles upload UI, job creation, job polling, and downloads.
- Calls `quote-api`.

### `quote-api`

- Hosts the Python HTTP API.
- Validates inputs and auth.
- Creates quote jobs.
- Returns job status.
- Returns signed download links or file metadata.

### `quote-worker`

- Runs Python background quote work.
- Performs AI reference-image analysis.
- Drafts quote basis.
- Handles quote-basis line edits.
- Generates XLSX output.
- Uploads generated files to durable storage.

### `redis`

- Add when jobs need queue durability across API restarts.
- Used by `quote-api` and `quote-worker`.

### `postgres`

- Use managed Supabase/Postgres first.
- Self-host only after backup, restore, upgrades, and monitoring are documented and tested.

## Request Flow

```text
Browser
  -> otc-frontend
  -> quote-api creates job
  -> quote-worker runs AI analysis and XLSX generation
  -> quote-worker stores generated XLSX
  -> otc-frontend polls quote-api
  -> user downloads quotation
```

## Storage

Temporary files:

```text
/data/swooshz/tmp
```

Generated outputs:

```text
/data/swooshz/output
```

Logs:

```text
/data/swooshz/logs
```

Production customer files must be copied to durable object storage. Do not rely on container-local storage for customer files.

## Environment Variables

Frontend:

```text
QUOTE_API_URL=https://api.example.com
```

API and worker:

```text
APP_MODE=deploy
AUTH_REQUIRED=true
SESSION_SECRET=<strong-random-secret>
QUOTE_OUTPUT_ROOT=/data/swooshz/output
QUOTE_TMP_ROOT=/data/swooshz/tmp
QUOTE_LOG_ROOT=/data/swooshz/logs
OPENAI_API_KEY=<openai-key-if-used>
OPENAI_DRAFT_MODEL=<model-id-or-app-alias>
OPENAI_BASIS_LINE_MODEL=<model-id-or-app-alias>
DEEPSEEK_API_KEY=<deepseek-key-if-used>
DEEPSEEK_MODEL=deepseek-v4-pro
```

## Python Dependencies

Install pinned Python dependencies before starting the API or worker:

```text
pip install --only-binary=:all: -r requirements.txt
```

`pypdfium2` is the preferred PDF page renderer for uploaded render decks, and
`Pillow` is used only as the local image encoder/compressor for the rendered
page bitmap. This avoids SaaS PDF conversion, avoids undeclared system
dependencies such as Poppler/`pdftoppm`, and keeps PDF rendering portable across
local and deployed runners. Rendered debug page copies stay under
`QUOTE_TMP_ROOT/pdf-pages/`.

The original PDF and rendered prompt images are still sent to OpenAI when the
operator explicitly runs remote AI quote analysis. Rendering itself does not
call an external server; provider analysis is the network boundary.

Set `DEEPSEEK_API_KEY` to use DeepSeek Pro for text-only work such as `Re`
selected-line edits, quote-basis answers/proposals, and pricing-reference import
normalization. `DEEPSEEK_MODEL` can be changed later to test another DeepSeek
model. Keep full PDF/image draft analysis on OpenAI unless a future provider
integration explicitly supports the same vision input contract. When a DeepSeek
text route fails or returns unusable JSON, the app may retry through OpenAI if
`OPENAI_API_KEY` is configured.

`pypdfium2` is preferred over PyMuPDF for deployment because pypdfium2/PDFium
uses permissive Apache-2.0/BSD-style licensing, while PyMuPDF/MuPDF is AGPL
unless a commercial license is used. Use binary-wheel-only installation so
deployment fails closed if a supported wheel is unavailable instead of running
an unreviewed source build. Keep dependencies pinned and update them
deliberately with advisory, release-note, and license review.

Storage and database variables:

```text
DATABASE_URL=<managed-database-url>
STORAGE_ENDPOINT=<object-storage-endpoint>
STORAGE_BUCKET=<bucket-name>
STORAGE_ACCESS_KEY=<storage-access-key>
STORAGE_SECRET_KEY=<storage-secret-key>
```

## Deployment Steps

1. Provision Hostinger VPS.
2. Install Coolify on Ubuntu LTS.
3. Point DNS to the VPS.
4. Create Coolify project.
5. Add `otc-frontend`.
6. Add `quote-api`.
7. Add `quote-worker`.
8. Add Redis if queue durability is needed.
9. Configure environment variables in Coolify.
10. Configure HTTPS through Coolify.
11. Configure persistent mounts for `/data/swooshz`.
12. Run a full quote from image upload to XLSX download.
13. Add external uptime monitoring.
14. Add off-server backup for any persistent VPS data.

## Production Gates

Do not treat the deployment as production-ready until these pass:

- Auth is enabled.
- HTTPS is enabled.
- AI keys are stored only in Coolify secrets/env.
- Customer uploads are size/type limited.
- Generated files are stored in durable storage.
- Logs do not expose secrets or customer file contents.
- Health checks pass for frontend and API.
- Worker restart does not lose queued jobs.
- Backup and restore procedure is tested.
- Server patching procedure is documented.
- Domain and DNS ownership are confirmed.

## Vercel / Render

Do not use Vercel + Render as the default deployment target.

Use it only if a later decision explicitly prioritizes managed frontend previews or managed backend hosting over single-server Coolify deployment. The optional Render example lives at `docs/examples/render.yaml` so the repository root does not imply Render is the production deployment path.

The current deploy/auth scaffold is guarded until a complete OIDC callback token exchange and claims-validation boundary is implemented. Do not expose a public deployment that routes users into the scaffolded callback; production access must stay blocked until the auth boundary is complete and tested.
