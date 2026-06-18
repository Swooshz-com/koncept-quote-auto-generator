# Sensitive Fixture Audit And Cleanup Preparation

This is a non-UI cleanup preparation note for KQAG committed fixtures and repo
packs. It does not authorize production deployment, auth, billing, public
exposure, or history rewrite work.

## Scope

Audited committed fixture/export-like surfaces:

- `profiles/koncept/**`
- `pricing-references/koncept-exhibition-quotation/**`
- `workspace-seeds/koncept-images-pte-ltd/asset-packs/**`
- `workspace-seeds/koncept-images-pte-ltd/workspace.json`
- `fixtures/samples/kent-group/**`
- docs XLSX/profile/pricing/export-like paths

The private maintainer exports are not repo fixtures and must stay outside the
repo. Do not copy their contents into tests, docs, logs, generated outputs, or
fixture replacements.

## Audit Summary

The scanner reports path and category only. It does not print matched values.
The current committed tree has no blocking scanner findings and has review
findings in these categories:

| Category | Paths |
| --- | --- |
| Company/workspace identity markers | `workspace-seeds/koncept-images-pte-ltd/workspace.json`, selected planning/checklist docs |
| Customer/sample markers | `fixtures/samples/kent-group/sample.json`, `fixtures/samples/kent-group/kent-group.pdf`, selected planning docs |
| Internal pricing fields | `pricing-references/koncept-exhibition-quotation/pricing-catalog.json`, `workspace-seeds/koncept-images-pte-ltd/asset-packs/pricing-references/koncept-workspace-pricing/pricing-catalog.json`, pricing import docs |
| Committed fixture media | `pricing-references/koncept-exhibition-quotation/pricing-catalog-images/**`, `workspace-seeds/koncept-images-pte-ltd/asset-packs/pricing-references/koncept-workspace-pricing/pricing-catalog-images/**` |
| Embedded logo/reference fields | `profiles/koncept/profile.json` |
| XLSX package internals requiring review | `profiles/koncept/quotation-layout.xlsx`, `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/koncept-workspace-template/quotation-layout.xlsx`, `docs/Quotation-Cost-Template-V1.1.xlsx`, `docs/examples/super-messy-pricing-reference.xlsx` |
| XLSX formulas/header-footer/shared strings/defined names/document properties | Same XLSX paths above, depending on package part |

XLSX inspection covered visible and hidden workbook metadata surfaces available
from the OOXML package: workbook sheet state, shared strings, defined names,
comments/notes, headers/footers, document properties, embedded media,
external-link package parts, formulas, and worksheet/sample row XML.

## Current-Tree Cleanup Recommendation

Remove, sanitise, or replace these committed packs in a later cleanup PR:

- `profiles/koncept/profile.json`
  - Replace with a deterministic synthetic quote-company profile.
  - Keep only schema/formatting coverage needed by tests.
  - Do not include real logo data, real signatory names, payment text, or
    private exported profile content.
- `profiles/koncept/quotation-layout.xlsx`
  - Replace with a deterministic generated XLSX layout fixture.
  - Ensure workbook properties, headers/footers, defined names, formulas,
    shared strings, and media are synthetic.
- `profiles/koncept/layout-rules.json`
  - Keep only generic layout behavior needed by generator tests.
  - Remove real-company-specific wording from rule names or comments if any are
    added later.
- `pricing-references/koncept-exhibition-quotation/**`
  - Replace with a deterministic synthetic pricing reference pack.
  - Remove committed catalog images unless a synthetic image fixture is required
    by an import/parser test.
  - Keep pricing fields only as synthetic test data.
- `workspace-seeds/koncept-images-pte-ltd/asset-packs/**`
  - Replace asset-pack contents with deterministic synthetic profile, template,
    and pricing fixtures.
  - Keep the workspace bridge behavior, but make fixture contents clearly fake.
- `fixtures/samples/kent-group/**`
  - Replace with a synthetic sample PDF/JSON pair or migrate tests to a
    deterministic generated sample.
  - Keep only the minimum sample needed for upload/PDF/browser smoke coverage.
- Docs XLSX examples under `docs/`
  - Keep only if they are clearly synthetic templates/examples.
  - Rebuild if package properties, comments, hidden sheets, formulas, media, or
    shared strings cannot be proven synthetic.

Files that are safe to keep after review:

- `pricing-references/import-cleanup-rules.json`
- scanner/test code that contains only synthetic trigger strings
- deterministic synthetic fixtures created from scratch for parser/generator
  tests

## Guardrails Added

- `.gitignore` now excludes local/private quote-company profile exports,
  private pricing uploads, generated quote outputs, screenshots, and local
  runtime data.
- `scripts/scan_sensitive_fixtures.py` scans committed fixture/export-like
  paths and returns non-zero for blocking categories.
- Known current-tree company/customer markers are allowlisted as review findings
  so this preparation PR can inventory them without silently accepting new
  ones. New unreviewed company/customer markers become blocking findings.
- `tests/test_sensitive_fixture_scan.py` verifies:
  - current committed fixture scan has no blocking findings
  - temporary private profile/payment markers fail with path/category-only
    output and no echoed values
  - unreviewed company/customer markers fail with path/category-only output

Run manually:

```powershell
python scripts\scan_sensitive_fixtures.py
```

Use review mode before a fixture replacement PR:

```powershell
python scripts\scan_sensitive_fixtures.py --fail-on-review
```

## Safe Fixture Plan

If removing committed packs breaks tests, replace them with deterministic
fixtures built from scratch:

1. Create a synthetic company identity that is not Koncept or a real customer.
2. Generate a small synthetic logo/media asset or use no media where the test
   does not require media.
3. Generate an XLSX layout fixture from deterministic code or a manually
   reviewed workbook with synthetic document properties and no hidden/private
   content.
4. Build a tiny pricing catalog with clearly fake sections, descriptions,
   costs, markups, aliases, and image fixtures.
5. Build a synthetic sample upload fixture and PDF if browser/PDF tests still
   require one.
6. Update tests to depend on generic behavior and schema contracts, not real
   company names, real customer names, real pricing, or real artwork.
7. Run the fixture scanner and generator/browser validation before merging.

Do not use private extracted local profile/pricing/template files as fixture
sources.

## Maintainer-Only History Rewrite Runbook

Do not run this in a normal feature PR. A repository owner must approve and
coordinate the rewrite explicitly.

1. Create a fresh mirror clone in a private maintenance directory:

   ```powershell
   git clone --mirror <repo-url> kqag-sensitive-cleanup.git
   cd kqag-sensitive-cleanup.git
   ```

2. Install and verify `git filter-repo` in the maintainer environment:

   ```powershell
   git filter-repo --version
   ```

3. Prepare a path removal/replacement list with path patterns only. Do not put
   sensitive values in the list.

4. Run filter-repo using sensitive data removal mode and path rules appropriate
   to the approved cleanup:

   ```powershell
   git filter-repo --sensitive-data-removal --path <path-to-remove> --invert-paths
   ```

   Repeat or use a paths file as needed for the approved set.

5. Recreate replacement synthetic fixtures in a normal clone after the rewrite,
   then open a reviewed PR for those replacement fixtures.

6. Stop for an explicit force-push approval gate. The approval must name the
   repository, branches/tags, and exact force-push operation.

7. If approved, force-push rewritten refs:

   ```powershell
   git push --force --all origin
   git push --force --tags origin
   ```

8. Warn collaborators that all existing clones, forks, PR branches, and cached
   refs may retain old objects. Require fresh clones after the rewrite.

9. Contact GitHub Support if old sensitive objects remain reachable through
   GitHub caches, pull request refs, forks, or repository network views.

10. Rotate any credential or payment secret that may ever have been committed,
    even if history was rewritten.
