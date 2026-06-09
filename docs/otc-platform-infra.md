# OTC AI Platform Infrastructure

## Recommended Setup

Use Vercel for the Next.js user-facing app, Render for the Python quote generation API, and Supabase later for auth, database, and generated file storage.

## Architecture

### Vercel Next.js

- Hosts the frontend.
- Hosts the user-facing app and dashboard.
- Handles UI state, quote-job creation, and polling.
- Calls the Render Python API for AI image analysis and quote generation.
- Shows waiting, running, completed, and failed states.

### Render Python FastAPI

- Hosts the Python quote service.
- Performs AI image analysis.
- Runs the existing quotation generator.
- Creates XLSX output.
- Stores job files temporarily during prototype phase.
- Later uploads generated files to Supabase Storage.

### Supabase Later

- Provides user authentication.
- Stores users, organizations, plans, entitlements, jobs, and usage events.
- Stores uploaded images and generated quotation files.
- Enables persistent downloads instead of relying on temporary Render disk.

## Request Flow

```text
User opens Vercel app
  -> uploads booth images and quote details
  -> Vercel creates quote job
  -> Vercel calls Render Python API
  -> Render drafts quote basis and generates XLSX
  -> Vercel polls job status
  -> user downloads generated quotation
```

## Render Free Behavior

Render free services may sleep when idle. The first request can take longer while the service wakes up.

Recommended user-facing behavior:

```text
Starting quote service...
This may take 30-60 seconds on the free plan.
```

Use polling instead of webhooks for quote generation.

## API Link

Vercel environment variables:

```text
QUOTE_API_URL=https://quote-api.onrender.com
QUOTE_API_SECRET=<shared-secret>
```

Render environment variables:

```text
QUOTE_API_SECRET=<same-shared-secret>
OPENAI_API_KEY=<openai-key>
OPENAI_DRAFT_MODEL=gpt-5.4-mini
OPENAI_BASIS_LINE_MODEL=gpt-5.4-nano
GEMINI_API_KEY=<gemini-key-if-used>
GEMINI_DRAFT_MODEL=gemini-flash-latest
GEMINI_BASIS_LINE_MODEL=gemini-3.1-flash-lite
```

Vercel calls Render with:

```text
Authorization: Bearer <QUOTE_API_SECRET>
```

## Future Production Flow

```text
Vercel Next.js
  -> starts quote job

Render Python API
  -> generates XLSX
  -> uploads output to Supabase Storage

Supabase
  -> stores job metadata and generated files

Vercel Next.js
  -> polls job status
  -> shows secure download link
```

## Why This Setup

- Vercel is best for Next.js.
- Render is better for Python and file generation.
- Supabase handles persistent auth, database, and storage later.
- The existing Python quote generator does not need to be rewritten.
- The architecture can grow into the full OTC AI solutions platform.
