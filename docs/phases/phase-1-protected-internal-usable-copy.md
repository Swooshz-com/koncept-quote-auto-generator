# Phase 1 - Protected Internal Usable Copy

## Goal

Let the internal team safely test and use the current quote generator while the
full platform continues to be built.

## Scope

- Prepare a protected staging path for controlled team testing. Local runs may
  support development, but they are not the internal usable copy.
- Keep access restricted behind an external protection gate before any public
  exposure.
- Make clear that temporary protection is not production auth.
- Keep quote generator behavior aligned with the current image-first quoting
  rules and `quotation.xlsx` as the customer-ready master output.
- Use only approved internal test data unless explicit current-turn approval
  names real customer data and the permitted operation.

## Non-Negotiables

- No hardcoded app password.
- No public exposure without an external protection gate.
- No claim that production auth is complete.
- No real customer data unless explicitly approved.
- No secrets committed to the repo.

## Exit Criteria

- Internal testers can reach a protected copy and run the current app safely.
- The protection boundary, limitations, and safe-use notes are documented.
- Smoke checks cover app health, upload-to-analysis readiness, quote-company
  profile loading, pricing reference selection, and Excel download where
  practical.
- Remaining platform-auth gaps are visible in PR notes and not described as
  complete.
