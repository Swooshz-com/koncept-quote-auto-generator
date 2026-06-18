# Sensitive Fixture Current-Tree Cleanup

This is a non-UI current-tree cleanup note for KQAG committed fixtures and repo
packs after PR #38. It does not authorize production deployment, auth, billing,
public exposure, or history rewrite work.

## Scope

Audited committed fixture/export-like surfaces and cleanup targets:

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
The cleaned committed tree is expected to have no blocking scanner findings.
Current-tree cleanup handled the review categories below by removal,
replacement, or documented sample-only retention:

| Category | Paths |
| --- | --- |
| Removed legacy bundled profile pack | `profiles/koncept/**` |
| Removed legacy bundled pricing pack | `pricing-references/koncept-exhibition-quotation/**` |
| Renamed synthetic workspace profile/template pack | `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/synthetic-exhibition-fixture-template/**` |
| Renamed synthetic workspace pricing pack | `workspace-seeds/koncept-images-pte-ltd/asset-packs/pricing-references/synthetic-exhibition-fixture-pricing/**` |
| Restored sample-only Kent upload fixture | `fixtures/samples/kent-group/**` |
| Workspace bridge identity retained | `workspace-seeds/koncept-images-pte-ltd/workspace.json` |
| XLSX package internals requiring review | `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/synthetic-exhibition-fixture-template/quotation-layout.xlsx` |
| XLSX formulas/header-footer/shared strings/defined names/document properties | Same XLSX paths above, depending on package part |

XLSX inspection covered visible and hidden workbook metadata surfaces available
from the OOXML package: workbook sheet state, shared strings, defined names,
comments/notes, headers/footers, document properties, embedded media,
external-link package parts, formulas, and worksheet/sample row XML.

## Current-Tree Cleanup Result

Removed committed legacy bundle paths:

- `profiles/koncept/**`
- `pricing-references/koncept-exhibition-quotation/**`

Kept the `koncept-images-pte-ltd` workspace bridge identity, but renamed the
active synthetic workspace asset-pack folders and ids:

- `koncept-workspace-template` to `synthetic-exhibition-fixture-template`
- `koncept-workspace-pricing` to `synthetic-exhibition-fixture-pricing`

The active committed runtime fixture packs are now:

- `workspace-seeds/koncept-images-pte-ltd/asset-packs/quotation-layouts/synthetic-exhibition-fixture-template/**`
- `workspace-seeds/koncept-images-pte-ltd/asset-packs/pricing-references/synthetic-exhibition-fixture-pricing/**`

The Kent sample was restored to its pre-PR #38 sample fixture state and remains
sample-only:

- `fixtures/samples/kent-group/sample.json`
- `fixtures/samples/kent-group/kent-group.pdf`

Files that remain safe to keep after review:

- scanner/test code that contains only synthetic trigger strings
- deterministic synthetic workspace seed fixtures created from scratch for
  parser, generator, import/export, logo-data-url, and smoke tests

## Guardrails Added

- `.gitignore` now excludes local/private quote-company profile exports,
  private pricing uploads, generated quote outputs, screenshots, and local
  runtime data.
- `scripts/scan_sensitive_fixtures.py` scans committed fixture/export-like
  paths and returns non-zero for blocking categories.
- Known safe synthetic fixture paths are reviewed in the scanner without
  printing matched values. New unreviewed company/customer/private/export-like
  markers become blocking findings.
- `tests/test_sensitive_fixture_scan.py` verifies:
  - current committed fixture scan has no blocking findings
  - temporary private profile/payment markers fail with path/category-only
    output and no echoed values
  - unreviewed company/customer markers fail with path/category-only output

Run manually:

```powershell
python scripts\scan_sensitive_fixtures.py
```

Use review mode before fixture changes:

```powershell
python scripts\scan_sensitive_fixtures.py --fail-on-review
```

## Safe Fixture Policy

Future fixture additions should stay deterministic and synthetic:

1. Use synthetic company identities that are not real customers.
2. Generate synthetic logo/media assets or omit media when tests do not require
   it.
3. Keep XLSX layout fixtures reviewed for synthetic document properties,
   headers/footers, defined names, formulas, shared strings, and media.
4. Keep pricing catalogs tiny and clearly fake, including sections,
   descriptions, costs, markups, aliases, and image fixtures.
5. Keep upload/PDF samples sample-only and detached from active runtime
   defaults.
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
