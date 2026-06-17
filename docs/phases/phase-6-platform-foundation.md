# Phase 6 - Real Platform Auth, Company, User, Role, Credit Ledger, And Billing Foundation

## Goal

Move from protected internal usage to the real Swooshz platform foundation.

## Scope

- The main Swooshz dashboard owns login, companies, users, roles, credits,
  entitlements, billing state, and product navigation.
- The quote generator consumes platform identity and company context instead of
  owning standalone local users.
- Koncept Images Pte Ltd local/staging bridge data migrates into the platform
  company model.
- Credits use a ledger model, not editable balance fields.
- Stripe/Supabase/OIDC work is implemented only with documented boundaries,
  tests, and secret handling.

## Required Platform Concepts

- companies
- users
- company memberships and roles
- product entitlements
- pricing reference packs and versions
- quote presets
- quote jobs and generated files
- AI usage events
- credit ledger
- billing customers and subscriptions

## Non-Negotiables

- Server-side entitlement checks before AI calls, uploads, quote generation, and
  downloads.
- RLS on exposed Supabase/Postgres tables when implemented.
- No service-role keys or billing secrets in frontend/static code.
- No frontend-calculated credit balance as an authorization boundary.
- No silent support impersonation.

## Exit Criteria

- Quote-generator entry validates platform user, company membership, role,
  product entitlement, and credit availability.
- Koncept Images Pte Ltd seed data has a documented migration path into real
  company records.
- Auth, billing, storage, and credit behavior have focused tests and clear
  rollback notes.
