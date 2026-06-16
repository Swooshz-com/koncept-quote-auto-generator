# MVP Implementation Plan

This is the canonical MVP implementation and infrastructure plan. Keep platform
architecture, app boundaries, auth, data ownership, billing, deployment, and
quote-generator product rules here instead of splitting them across separate
planning docs.

## Goal

Ship the Swooshz product platform and quote generator for controlled MVP testing
on Hostinger VPS + Coolify without painting ourselves into a single-company or
local-only corner.

The quote generator is one product module inside the Swooshz platform. The main
authenticated dashboard owns companies, users, credits, entitlements, billing
state, and cross-product navigation. The quote generator must consume that
platform identity and company context instead of building a separate user-account
system.

The pricing reference is the customer's owned offer catalog. Quote-basis
generation must use pricing-reference rows as aggressively as possible, and use
AI custom/proposal rows only for genuinely unmatched optional suggestions that
the company may choose to source or ignore.

## Product Direction

- Keep the public marketing site simple and mostly informational.
- Put the authenticated platform dashboard at a separate app origin, for
  example `app.swooshz.com`.
- Use one identity layer for all Swooshz products. Company admins and users log
  into the main dashboard first, then enter the quote generator and other
  product modules through the same platform session or a short-lived internal
  session handoff.
- Build company administration in the main dashboard first: company profile,
  company users, roles, monthly credits, product entitlements, and billing view.
- Treat bundled customer data as seed/demo data only, not as a global app
  default for every future customer.
- Keep pricing references, presets, quote jobs, uploaded files, generated files,
  usage logs, and billing status scoped to a company/account.
- Keep the whitelabelled external solution as a separate product integration.
  The Swooshz platform owns the customer account, entitlements, and launch link;
  the external provider owns its own runtime. Use SSO, signed handoff, or
  provider-supported user provisioning rather than mixing its infrastructure into
  the quote-generator backend.

## MVP Source Of Truth

Use managed products where they remove custom security work without creating
expensive MVP commitments:

- Supabase is the MVP source of truth for auth, company/user relational data,
  memberships, roles, RLS-protected app data, audit events, AI usage events,
  credit ledgers, entitlements, and possibly storage.
- Stripe is the billing source of truth. Mirror only the subscription state,
  plan, period end, cancellation status, customer id, subscription id, and
  entitlement limits into Supabase/Postgres for fast authorization and dashboards.
- HubSpot is optional CRM sync only. It may mirror contacts, companies, deals,
  lifecycle stage, and owner notes, but it must not become the source of truth
  for auth, roles, credits, billing, pricing references, quote jobs, files, or
  entitlements.
- Hostinger VPS + Coolify remains the first app/runtime deployment target.
  Keep managed Supabase/Postgres available for customer data until self-hosted
  backup, restore, upgrades, and monitoring are documented and tested.

Free-to-start guidance:

- Supabase is suitable for MVP start, but verify current plan limits before
  implementation because Supabase pricing and quotas can change.
- Stripe has no app subscription requirement for basic setup, but live payments
  and Stripe Billing create transaction or billing-volume fees.
- HubSpot free CRM can be useful for early sales tracking, but sync should stay
  optional so paid HubSpot features do not block the product MVP.
- AI provider usage, VPS hosting, domains, email, and storage overages are not
  free. Credit caps and usage logs are required before public usage.

## Roles

Use three product roles for MVP:

- `admin`: manage company settings, users, billing view, pricing references,
  company presets, quote generation, and downloads.
- `management`: manage pricing references, company presets, quote generation,
  and downloads, but not billing ownership or destructive company-level actions.
- `operator`: generate quotes, use approved pricing references, save/load own
  presets, and download generated quote files.

Do not add a viewer role for MVP unless a customer explicitly needs read-only
access later.

## Company Accounts And Credits

Company admins manage subaccounts from the main dashboard. The quote generator
does not create standalone local users.

MVP account flow:

1. Swooshz owner/admin creates or approves a company.
2. Stripe Checkout or manual internal setup creates the company's billing
   relationship.
3. Stripe webhooks update mirrored subscription and entitlement rows.
4. Swooshz owner/admin can grant monthly credits or plan overrides.
5. Company admins invite users and assign roles.
6. Company admins can allocate user-level usage caps or let all users draw from
   the shared company pool.
7. Quote generator AI calls consume credits server-side and write usage events.

Credit rules:

- Use a ledger model, not an editable balance field. Store grants, monthly
  allocations, manual adjustments, usage debits, refunds, expiries, and
  reversals as separate rows.
- Entitlement checks must happen server-side before AI calls, uploads, quote
  generation, downloads, and external-product launch links.
- Monthly credit allocation should be derived from Stripe plan limits plus any
  Swooshz owner override.
- Do not let frontend flags, hidden form fields, or client-calculated balances
  decide access.

## Catalog-First Quote Basis

Rules to preserve:

- The pricing reference is the company's actual sell/rent/service catalog.
- AI must first ask whether any pricing-reference row can safely represent the
  observed item.
- Catalog-backed lines must preserve exact pricing-reference wording.
- AI proposal/custom lines are only for truly weird, special, or optional
  suggested scope outside the pricing reference.
- Avoid broad one-lot custom rows when pricing-reference rows exist for the
  observed category, material, equipment, service, or connection scope.
- Do not force unsafe matches. Distinguishing attributes such as object family,
  shape, colour, finish, material, size, capacity, and mounting still matter.
- Pricing-reference imports must save a deterministic metadata baseline
  (`match_terms` and `object_families`) from row wording, aliases, remarks,
  units, sections, and visuals. This baseline must be generic across industries,
  not tuned to furniture, exhibitions, or any single sample catalog.
- After the deterministic baseline is saved, run AI metadata enrichment when a
  provider is configured to improve generic synonyms and object-family coverage
  for unfamiliar industries. This enrichment must be post-save, optional/manual,
  or async; it must never block Save.
- AI metadata enrichment output must be validated row-by-row, must not rewrite
  descriptions, prices, units, ids, sections, or remarks, and must fall back to
  deterministic metadata if the provider fails or returns weak output. Every
  enrichment attempt must be logged as an AI call with the same privacy-safe
  metadata rules as other AI features.

## Data Model

Add durable storage before public MVP. Supabase/Postgres should own these
relational records:

- `companies`
- `users`
- `company_memberships`
- `company_product_entitlements`
- `external_product_connections`
- `pricing_reference_packs`
- `pricing_reference_versions`
- `quote_presets`
- `quote_jobs`
- `uploaded_files`
- `generated_files`
- `ai_usage_events`
- `credit_ledger`
- `user_credit_limits`
- `audit_events`
- `billing_customers`
- `subscriptions`

Recommended storage split:

- Postgres for metadata, memberships, permissions, billing state, usage events,
  audit events, presets, and pricing-reference records.
- Supabase Storage or S3-compatible durable object storage for uploaded
  PDFs/images and generated XLSX files.
- Redis only when background job durability is needed across API restarts.

Supabase/RLS rules:

- Enable RLS on every exposed table.
- Do not authorize from user-editable `user_metadata`; store authorization data
  in app-controlled tables or app metadata.
- Every company-scoped table must include `company_id` and have policies that
  restrict access to active memberships for that company.
- Owner/internal support views must avoid silent impersonation. Use support
  metadata and error references, not customer content.
- Keep service-role keys and other privileged credentials out of frontend code.

## Presets

Support these scopes:

- System default preset.
- Company default preset.
- Company named presets.
- User default preset.
- User named presets.

Each preset should store owner scope, version, last editor, and timestamps.
Operators can save their own presets. Admin/management can save company presets
and set company defaults.

## Auth And Permissions

MVP must replace local `USER_TYPE` simulation in deploy mode with Supabase Auth
users, platform sessions, and company memberships.

Minimum auth behavior:

- Login/logout.
- Session cookies with `Secure`, `HttpOnly`, and `SameSite=Lax` where a server
  session is used.
- Server-side permission checks on every sensitive endpoint.
- Company boundary check on every job, file, pricing reference, preset, and
  settings request.
- Admin can add/invite/deactivate users and change roles.
- Deactivated users cannot create jobs or download company files.
- Quote generator entry must validate the platform user, company membership,
  product entitlement, role, and credit availability.

Current OIDC scaffolding is not enough for public deployment until token
exchange, claims validation, user lookup, membership lookup, and session
creation are complete and tested.

## AI Usage Logging And Abuse Control

Track AI usage to stop abuse while keeping PDPA/GDPR data minimization.

Log every AI call attempt, not only errors. This includes successful calls,
failed calls, retries, fallbacks, deterministic non-AI bypasses where relevant,
and blocked calls that did not reach a provider.

Log per AI request attempt:

- `company_id`
- `user_id`
- job id when available
- provider and model
- feature/mode, such as draft analysis, Re, pricing import
- request start/end timestamps
- duration
- image/PDF page counts
- retry/fallback chain
- approximate input/output token counts when available
- estimated cost when available
- status, error category, retry count, and error reference
- rate-limit decision
- support/feedback correlation id so complaints can be traced to the relevant
  AI attempt without storing the prompt or generated content

Do not store raw prompts, raw images, raw PDFs, full generated text, API keys,
cookies, Authorization headers, generated quote contents, quote-basis text,
pricing-reference contents, or long customer text in usage logs.

Abuse controls:

- Per-user and per-company daily/monthly AI spend caps.
- Per-user and per-company request rate limits.
- Admin override to suspend a user or company.
- Soft warning threshold before hard blocking.
- Internal admin dashboard showing usage by company, user, model, feature, and
  time period.
- Audit event whenever a user is blocked, unblocked, role-changed, or limit
  changed.

## Billing

Use Stripe Billing + Checkout for MVP subscriptions and Stripe Customer Portal
for billing self-service:

- Stripe Checkout: https://docs.stripe.com/payments/checkout
- Stripe subscriptions: https://docs.stripe.com/billing/subscriptions/overview
- Stripe Customer Portal: https://docs.stripe.com/customer-management

Recommended MVP billing flow:

1. Public marketing/pricing page links to Stripe Checkout.
2. Stripe webhook confirms subscription status.
3. App creates or updates the company billing record.
4. Entitlements, monthly credits, and plan limits are stored server-side.
5. App access, product access, SSO handoff, and AI caps depend on server-side
   billing and entitlement state, not frontend flags.
6. Admin users can open Stripe Customer Portal from the app to manage billing.

Use Stripe as billing source of truth, but mirror subscription status, plan,
period end, cancellation status, and entitlement limits into Postgres for fast
authorization and dashboards.

Do not build manual subscription renewal loops with raw PaymentIntents. Use
Stripe Billing, Checkout Sessions in subscription mode, Stripe Prices, webhooks,
and Customer Portal.

## Product Modules And External SSO

The main dashboard is the product shell. Product modules are launched from it.

MVP modules:

- Quote generator: first-party Swooshz module, using platform auth, company
  context, credits, pricing references, quote jobs, and generated files.
- Whitelabelled external product: separate provider/runtime, launched from
  Swooshz only after server-side entitlement checks.

External product integration rules:

- Store external account/user ids in `external_product_connections`.
- Prefer provider-supported SSO or signed launch links.
- Do not duplicate passwords or store third-party user credentials.
- Log launch attempts as audit events without storing sensitive external tokens.
- Keep the external product's infra separate unless a future migration explicitly
  brings it under Swooshz ownership.

## Swooshz Owner Dashboard

Build a separate owner/admin area inside the authenticated main dashboard, not
visible to customer companies.

MVP owner dashboard:

- Companies list.
- Subscription status and plan.
- Product entitlements.
- Monthly credit allocation and remaining credits.
- User count.
- AI usage and estimated cost.
- Job count and failure rate.
- Abuse flags and blocked users.
- Pricing reference count/version age.
- Last activity.
- Impersonation-free support links: view metadata and error references, never
  silently act as customer users.

This dashboard belongs inside the authenticated app/admin area, not the public
marketing pages. The marketing site can link to login, signup, pricing, and docs.

## Target Infrastructure

Use Hostinger VPS + Coolify for the first hosted app deployment.

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

Do not expose databases, caches, or admin ports publicly. Keep managed
Supabase/Postgres as the first production data store. Self-host Postgres only
after backup, restore, upgrades, monitoring, and security operations are tested.

## Coolify Services

Use the Hostinger/Coolify setup skill rules:

- Inspect and record evidence before changing the VPS.
- Do not expose databases, caches, or admin ports publicly.
- Store secrets only in Coolify/server-managed configuration.
- Require owner approval before DNS, firewall, env var, reboot, destructive, or
  production deploy actions.
- Record deployment evidence and rollback notes under
  `docs/hostinger-coolify/`.

MVP app services:

- `swooshz-dashboard`: authenticated main dashboard, company admin, users,
  credits, entitlements, billing portal links, and product launch navigation.
- `quote-api`: Python HTTP API for quote jobs, pricing references, uploads,
  polling, downloads, and server-side authorization.
- `quote-worker`: Python background quote work, remote AI calls, quote-basis
  edits, XLSX generation, and durable file uploads.
- `redis`: add only when jobs need queue durability across API restarts.
- `postgres`: use managed Supabase/Postgres first.
- `storage`: use Supabase Storage or S3-compatible durable storage for uploads
  and generated files.

Request flow:

```text
Browser
  -> swooshz-dashboard authenticates user and company context
  -> quote-api validates membership, role, entitlement, and credits
  -> quote-api creates quote job
  -> quote-worker runs AI analysis and XLSX generation
  -> quote-worker stores generated XLSX in durable storage
  -> dashboard or quote UI polls quote-api
  -> user downloads quotation through an authorized signed link
```

Local container paths:

```text
QUOTE_TMP_ROOT=/data/swooshz/tmp
QUOTE_OUTPUT_ROOT=/data/swooshz/output
QUOTE_LOG_ROOT=/data/swooshz/logs
```

Production customer files must be copied to durable object storage. Do not rely
on container-local storage for customer files.

Frontend env:

```text
QUOTE_API_URL=https://api.example.com
```

API and worker env:

```text
APP_MODE=deploy
AUTH_REQUIRED=true
SESSION_SECRET=<strong-random-secret>
QUOTE_OUTPUT_ROOT=/data/swooshz/output
QUOTE_TMP_ROOT=/data/swooshz/tmp
QUOTE_LOG_ROOT=/data/swooshz/logs
DATABASE_URL=<managed-supabase-postgres-url>
SUPABASE_URL=<supabase-project-url>
SUPABASE_PUBLISHABLE_KEY=<publishable-key-if-used-by-server-rendered-client>
SUPABASE_SERVICE_ROLE_KEY=<server-only-service-role-key>
STORAGE_ENDPOINT=<object-storage-endpoint>
STORAGE_BUCKET=<bucket-name>
STORAGE_ACCESS_KEY=<storage-access-key>
STORAGE_SECRET_KEY=<storage-secret-key>
STRIPE_SECRET_KEY=<stripe-secret-key>
STRIPE_WEBHOOK_SECRET=<stripe-webhook-secret>
OPENAI_API_KEY=<openai-key-if-used>
OPENAI_DRAFT_MODEL=<model-id-or-app-alias>
OPENAI_DRAFT_REASONING_EFFORT=high
OPENAI_DRAFT_HIGH_QUALITY_REASONING_EFFORT=xhigh
OPENAI_BASIS_LINE_MODEL=<model-id-or-app-alias>
DEEPSEEK_API_KEY=<deepseek-key-if-used>
DEEPSEEK_MODEL=deepseek-v4-pro
DEEPSEEK_BASIS_LINE_MODEL=deepseek-v4-flash
DEEPSEEK_BASIS_ANSWER_MODEL=deepseek-v4-flash
DEEPSEEK_BASIS_PROPOSAL_MODEL=deepseek-v4-pro
DEEPSEEK_PRICING_IMPORT_MODEL=deepseek-v4-pro
DEEPSEEK_PRICING_METADATA_MODEL=deepseek-v4-flash
```

Never put service-role keys, Stripe secret keys, webhook secrets, AI provider
keys, or storage secrets in static/frontend code.

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

Set `DEEPSEEK_API_KEY` to use DeepSeek for text-only work. Low-risk label-maker
routes can use Flash by default (`DEEPSEEK_BASIS_LINE_MODEL`,
`DEEPSEEK_BASIS_ANSWER_MODEL`, and `DEEPSEEK_PRICING_METADATA_MODEL`), while
whole-basis proposals and messy pricing import normalization stay on Pro
(`DEEPSEEK_BASIS_PROPOSAL_MODEL` and `DEEPSEEK_PRICING_IMPORT_MODEL`).
`DEEPSEEK_MODEL` remains the legacy/global DeepSeek fallback model. A custom
non-Pro value still acts as a global override, but `DEEPSEEK_MODEL=deepseek-v4-pro`
does not suppress Flash route defaults.
Keep full PDF/image draft analysis on OpenAI unless a future provider
integration explicitly supports the same vision input contract. When a DeepSeek
text route fails or returns unusable JSON, the app may retry the next configured
DeepSeek model and then OpenAI if `OPENAI_API_KEY` is configured.

`pypdfium2` is preferred over PyMuPDF for deployment because pypdfium2/PDFium
uses permissive Apache-2.0/BSD-style licensing, while PyMuPDF/MuPDF is AGPL
unless a commercial license is used. Use binary-wheel-only installation so
deployment fails closed if a supported wheel is unavailable instead of running
an unreviewed source build. Keep dependencies pinned and update them
deliberately with advisory, release-note, and license review.

## Deployment Steps

1. Provision Hostinger VPS.
2. Install Coolify on Ubuntu LTS.
3. Point DNS to the VPS.
4. Create Coolify project.
5. Add `swooshz-dashboard`.
6. Add `quote-api`.
7. Add `quote-worker`.
8. Connect managed Supabase/Postgres.
9. Configure Supabase Auth, RLS, storage buckets, and server-only credentials.
10. Configure Stripe products, prices, Checkout, Customer Portal, and webhooks.
11. Add Redis only if queue durability is needed.
12. Configure environment variables in Coolify.
13. Configure HTTPS through Coolify.
14. Configure persistent mounts for `/data/swooshz`.
15. Run a full quote from image upload to XLSX download.
16. Run company/user/role/credit entitlement smoke tests.
17. Run Stripe webhook and Customer Portal smoke tests in test mode.
18. Run external-product launch-link smoke tests if the whitelabelled product is
    enabled.
19. Add external uptime monitoring.
20. Add off-server backup for any persistent VPS data.

Production gates:

- Auth complete.
- HTTPS complete.
- Company isolation tested.
- Supabase RLS tested.
- User invite/deactivation and role changes tested.
- AI usage caps active.
- Credit ledger and monthly allocation tested.
- Billing webhook verified.
- Stripe Customer Portal tested.
- Whitelabel product SSO/handoff tested if enabled.
- Backups and restore procedure documented.
- Customer uploads are size/type limited.
- Generated files are stored in durable storage.
- Logs do not expose secrets or customer file contents.
- Health checks pass for dashboard, API, and worker.
- Worker restart does not lose queued jobs once Redis/queue durability is
  introduced.
- Server patching procedure is documented.
- Domain and DNS ownership are confirmed.
- Full upload-to-XLSX smoke test passes.

## Non-Coolify Hosting

Do not use Vercel + Render as the default deployment target, and do not keep
stale provider-specific examples in the repo while Hostinger VPS + Coolify is
the selected MVP deployment path.

Add a provider-specific deployment example only after a later decision
explicitly prioritizes that provider over single-server Coolify deployment, and
keep that example aligned with the current auth, storage, dependency, and
secret-handling requirements.

The current deploy/auth scaffold is guarded until a complete OIDC callback token
exchange and claims-validation boundary is implemented. Do not expose a public
deployment that routes users into the scaffolded callback; production access
must stay blocked until the auth boundary is complete and tested.

## Antigravity UI Prompt

Use this prompt for UI polish:

```text
You are polishing the existing Swooshz quote generator UI.

Important: do not redesign the app. The current general feel, layout direction,
visual language, and workflow are good. Keep the existing product identity and
only improve clarity, spacing, consistency, responsiveness, and interaction
details.

Preserve:
- Existing workflow: Images, Customer, Quote Company, Basis, Output.
- Existing overall look and feel.
- Existing app structure, page density, and operational-tool vibe.
- Existing element IDs, data attributes, and app behavior.
- Current visual hierarchy unless there is a clear usability issue.

Improve only:
- Alignment, spacing, sizing, and scanability.
- Button consistency and icon affordances.
- Modal and side-panel polish.
- Quote basis row readability.
- Pricing-reference-backed vs AI proposal/custom row clarity.
- Empty/loading/error/review states.
- Mobile/tablet responsiveness.
- Text wrapping so labels/buttons never overlap or clip.
- Small visual inconsistencies that make the UI feel unfinished.

Product rule:
The pricing reference is the customer's owned offer catalog. Catalog-backed rows
should feel like the normal/default outcome. AI proposal/custom rows should feel
like exceptions or optional suggestions needing extra attention.

Do not:
- Create a landing page.
- Add hero sections, marketing blocks, decorative gradients, blobs, or big visual
  flourishes.
- Change the core workflow.
- Replace the current visual style with a new design system.
- Remove security, permission, upload-limit, pricing-reference, or quote-generation
  guards.
- Hide important review states just to make the UI cleaner.

Deliver:
- Minimal CSS/layout/component polish.
- Before/after screenshots for desktop and mobile.
- A short summary of exactly what was changed and why.
```
