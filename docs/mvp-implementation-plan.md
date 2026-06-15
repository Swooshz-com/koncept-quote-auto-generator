# MVP Implementation Plan

## Goal

Ship the quote generator for controlled MVP testing on Hostinger VPS + Coolify
without painting ourselves into a single-company/local-only corner.

The pricing reference is the customer's owned offer catalog. Quote-basis
generation must use pricing-reference rows as aggressively as possible, and use
AI custom/proposal rows only for genuinely unmatched optional suggestions that
the company may choose to source or ignore.

## Product Direction

- Keep the public marketing site simple and mostly informational.
- Put the authenticated product at a separate app origin, for example
  `app.swooshz.com`.
- Use one identity layer for all Swooshz products, then let users enter the
  quote generator through SSO/session handoff.
- Treat bundled customer data as seed/demo data only, not as a global app
  default for every future customer.
- Keep pricing references, presets, quote jobs, uploaded files, generated files,
  usage logs, and billing status scoped to a company/account.

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

## Data Model

Add durable storage before public MVP:

- `companies`
- `users`
- `company_memberships`
- `pricing_reference_packs`
- `pricing_reference_versions`
- `quote_presets`
- `quote_jobs`
- `uploaded_files`
- `generated_files`
- `ai_usage_events`
- `audit_events`
- `billing_customers`
- `subscriptions`

Recommended storage split:

- Postgres for metadata, memberships, permissions, billing state, usage events,
  audit events, presets, and pricing-reference records.
- Durable object/file storage for uploaded PDFs/images and generated XLSX
  files.
- Redis only when background job durability is needed across API restarts.

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

MVP must replace local `USER_TYPE` simulation in deploy mode with authenticated
users and company memberships.

Minimum auth behavior:

- Login/logout.
- Session cookies with `Secure`, `HttpOnly`, and `SameSite=Lax`.
- Server-side permission checks on every sensitive endpoint.
- Company boundary check on every job, file, pricing reference, preset, and
  settings request.
- Admin can add/invite/deactivate users and change roles.
- Deactivated users cannot create jobs or download company files.

Current OIDC scaffolding is not enough for public deployment until token
exchange, claims validation, user lookup, membership lookup, and session
creation are complete and tested.

## AI Usage Logging And Abuse Control

Track AI usage to stop abuse while keeping PDPA/GDPR data minimization.

Log per AI request:

- `company_id`
- `user_id`
- job id
- provider and model
- feature/mode, such as draft analysis, Re, pricing import
- request start/end timestamps
- image/PDF page counts
- approximate input/output token counts when available
- estimated cost when available
- status, retry count, and error reference
- rate-limit decision

Do not store raw prompts, raw images, raw PDFs, full generated text, API keys,
cookies, Authorization headers, or long customer text in usage logs.

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
4. Entitlements/plan limits are stored server-side.
5. App access and AI caps depend on server-side billing state, not frontend
   flags.
6. Admin users can open Stripe Customer Portal from the app to manage billing.

Use Stripe as billing source of truth, but mirror subscription status, plan,
period end, cancellation status, and entitlement limits into Postgres for fast
authorization and dashboards.

## Swooshz Owner Dashboard

Build a separate owner/admin area, not visible to customer companies.

MVP owner dashboard:

- Companies list.
- Subscription status and plan.
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

## Coolify Deployment

Use the Hostinger/Coolify setup skill rules:

- Inspect and record evidence before changing the VPS.
- Do not expose databases, caches, or admin ports publicly.
- Store secrets only in Coolify/server-managed configuration.
- Require owner approval before DNS, firewall, env var, reboot, destructive, or
  production deploy actions.
- Record deployment evidence and rollback notes under
  `docs/hostinger-coolify/`.

MVP services:

- `quote-web` or combined MVP app service.
- Postgres, preferably managed first unless self-hosted backup/restore is tested.
- Redis only when queue durability is required.
- Durable storage for uploads and generated files.

Production gates:

- Auth complete.
- HTTPS complete.
- Company isolation tested.
- AI usage caps active.
- Billing webhook verified.
- Backups and restore procedure documented.
- Full upload-to-XLSX smoke test passes.

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
The pricing reference is the customer’s owned offer catalog. Catalog-backed rows
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
