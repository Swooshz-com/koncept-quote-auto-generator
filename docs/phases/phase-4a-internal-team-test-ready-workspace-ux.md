# Phase 4A - Internal Team Test-Ready Workspace UX

Phase 4A keeps the quote generator local/staging friendly for Koncept Images Pte
Ltd internal testing. It is not a deployment, auth, billing, ecommerce, or
platform-company milestone.

## Main Quote Flow

The main flow is for daily quote generation:

- Upload booth render or reference files.
- Fill Customer details.
- Select the quote pricing reference.
- Select and load the quote-company profile.
- Review Quote Basis.
- Review Output rows and download the Excel workbook.

Saved-profile management controls are intentionally not in the main
Quote Company step. The main step only selects the current quote-company
profile and loads it into the draft.

## Workspace Settings

Use **Settings** for workspace management tasks:

- Check the active company and workspace ID.
- Check the active quote-company profile source.
- Check the active pricing reference and source.
- Check the active quotation template/layout source.
- Import, export, save, or delete saved quote-company profiles.
- Review, import, export, edit, save, or delete repo-backed pricing references
  where the current local permissions allow it.

The active Koncept workspace is:

- Company: `Koncept Images Pte Ltd`
- Workspace ID: `koncept-images-pte-ltd`
- Workspace quotation template pack: `koncept-workspace-template`
- Workspace pricing reference pack: `koncept-workspace-pricing`

Workspace-owned packs are the preferred local/staging sources. The bundled
Koncept profile, bundled pricing reference, and sample fixtures remain fallback
or test fixtures until a later cleanup gate removes them.

## Verification

For a local/staging check:

- Open the app and confirm the workspace strip shows Koncept Images Pte Ltd.
- Open Settings and confirm the workspace ID, active profile, active pricing
  reference, and template source.
- Import a quote-company profile JSON and confirm its `logo_data_url` restores
  into the draft profile.
- Generate a test workbook from uploaded reference images and confirm the
  resulting `quotation.xlsx` contains the header logo.
- Confirm destructive saved-profile and pricing-reference controls are visible
  only in Settings, not in the normal quote flow.

## Later Platform Work

Do not add production auth, Supabase, Stripe, Hostinger/Coolify deployment,
customer accounts, credit ledger, billing, or role-management behavior in this
phase. Those remain later platform/deployment phases in the MVP plan.
