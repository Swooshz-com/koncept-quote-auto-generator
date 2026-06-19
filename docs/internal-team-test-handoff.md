# Internal Team Test Handoff

This runbook covers local/internal testing of the KQAG/SAQG quote generator
module only. The expected team path is a clean runtime workspace, private
runtime imports, and Excel output verification. Do not put private quote data in
the repository.

## Module Scope

KQAG owns quote-specific workflow and settings:

- booth/render image intake
- quote-company profile import/export
- pricing-reference import, review, save, edit, and selection
- quote basis review
- pricing review
- `quotation.xlsx` generation

The main Swooshz platform will eventually own login, accounts, billing, credits,
company access, product navigation, and app entitlement. Those platform concerns
are out of scope for this repository unless a future task explicitly moves this
module into that platform boundary.

## Local-Only Asset Rules

- Keep private files outside the repo.
- Do not commit private profile JSON, pricing XLSX, logo data, workbook rows,
  bank details, customer data, screenshots with private data, or generated
  customer quotes.
- Do not copy private files into `docs/`, `fixtures/`, `profiles/`,
  `pricing-references/`, `tests/fixtures/`, or any tracked fixture path.
- Use placeholders when reporting file locations, such as
  `<private-profile-json-outside-repo>` and
  `<private-pricing-xlsx-outside-repo>`.
- Runtime/local storage should stay under ignored paths selected with env vars
  such as `QUOTE_DATA_ROOT=<local-runtime-data-root>`,
  `QUOTE_TMP_ROOT=<local-runtime-temp-root>`,
  `QUOTE_OUTPUT_ROOT=<local-runtime-output-root>`,
  `QUOTE_LOG_ROOT=<local-runtime-log-root>`, and
  `KQAG_LOCAL_PRICING_REFERENCES_ROOT=<local-runtime-pricing-root>`.
- Synthetic automated-test fixtures live under
  `tests/fixtures/quote-generator/` and are not user/team-facing data.

## Clean Runtime Checklist

Use a fresh runtime workspace before each internal team test run.

1. Stop the local webapp if it is running.
2. Confirm `git status --short` is clean.
3. Choose runtime roots outside the repo and set the env vars listed above.
4. Restart the local webapp.
5. Manually import `<private-profile-json-outside-repo>`.
6. Manually import `<private-pricing-xlsx-outside-repo>`.
7. Generate and export a quote through the normal flow.
8. Confirm `git status --short` remains clean.

## Fresh-Pull Setup

1. Pull the latest `main`.
2. Keep private profile/pricing files outside the repo.
3. Set runtime roots outside the repo if needed.
4. Start the local webapp:

   ```powershell
   python webapp/server.py --host 127.0.0.1 --port 8765
   ```

5. Open `http://127.0.0.1:8765/`.

## Import Profile And Pricing

For the quote-company profile:

1. Open the Quote Company step.
2. Use the profile import control.
3. Select `<private-profile-json-outside-repo>`.
4. Confirm the imported profile is selectable.

For the pricing reference:

1. Open Pricing Reference Settings.
2. Go to Import.
3. Select `<private-pricing-xlsx-outside-repo>`.
4. Review the imported rows.
5. Save the pricing reference.
6. Return to Manage and confirm the saved reference is selectable.

Expected result: private runtime imports remain in ignored local/runtime storage
and become available to the normal quote flow.

## Generate And Verify A Quote

1. Start from a fresh quote.
2. Upload booth render images or other approved internal test files.
3. Complete Customer and Quote Company details.
4. Confirm the imported private quote-company profile is selected.
5. Confirm the imported private pricing reference is selected.
6. Run AI analysis.
7. Review and confirm the Quote Basis.
8. Resolve pricing review rows.
9. Click Download Excel.

Verify the exported `quotation.xlsx`:

- opens in Excel without a repair prompt
- shows the expected logo if the imported profile includes one
- shows GST/company profile values correctly
- uses `sqm` for square-metre quantities
- does not expose internal cost, supplier notes, or hidden private fields unless
  intentionally included by the internal test owner
- keeps nonnumeric unit price overrides safe as `???`
- does not include an unexpected trailing blank print page

The app does not generate PDFs by default. If an internal PDF is needed, create
it from the verified XLSX in Excel.

## Troubleshooting

If the imported pricing reference is not selectable:

1. Confirm the save step completed in Pricing Reference Settings.
2. Reopen Pricing Reference Settings -> Manage.
3. Hard refresh the page if the server was restarted.
4. Confirm `git status --short` did not show private files under tracked paths.

If a browser tab was open before the local server restarted, hard refresh the
tab and retry the operation.

If Excel shows a repair prompt, report only the issue category, browser, branch,
and step. Do not paste private workbook XML, rows, values, logo Base64, customer
data, or local file paths into GitHub.

## Smoke Checklist

- [ ] Clean runtime workspace selected before import.
- [ ] Profile import succeeds.
- [ ] Pricing import succeeds.
- [ ] Imported pricing reference is selectable.
- [ ] Quote generation succeeds through the normal flow.
- [ ] XLSX opens without Excel repair prompt.
- [ ] Logo appears when the private profile includes a logo.
- [ ] GST/company profile values appear correctly.
- [ ] Nonnumeric unit price override displays as `???`.
- [ ] `git status --short` stays clean.
