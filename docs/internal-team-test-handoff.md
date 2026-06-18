# Internal Team Test Handoff

This runbook is for local/internal Koncept Images team testing of the current
KQAG/SAQG quote flow. It is meant to make the private profile and pricing
import smoke repeatable without putting private assets into the repo.

## Scope

This phase is:

- Local/internal testing of the current quote workflow.
- Private asset import into local runtime storage only.
- Verification that the normal quote-facing flow uses the selected company
  profile and selected pricing reference.
- Excel output verification before broader platform work.

This phase is not:

- Production deployment.
- Public exposure.
- Login/auth rollout.
- Customer accounts.
- Database, Supabase, Stripe, billing, Hostinger, or Coolify work.
- A replacement for later platform/company persistence.

## Local-Only Asset Rules

- Keep private files outside the repo.
- Do not commit private profile JSON, pricing XLSX, logo data, workbook rows,
  bank details, customer data, or exported customer quotes.
- Do not copy private files into `docs/`, `fixtures/`, `profiles/`,
  `pricing-references/`, `workspace-seeds/`, or any tracked fixture path.
- Use placeholders when reporting file locations, such as
  `<private-profile-json-outside-repo>` and
  `<private-pricing-xlsx-outside-repo>`.
- Runtime/local storage may use env vars such as
  `QUOTE_DATA_ROOT=<local-runtime-data-root>` and
  `KQAG_LOCAL_PRICING_REFERENCES_ROOT=<local-runtime-pricing-root>`.
- Keep generated outputs in ignored runtime/output folders only.

## Fresh-Pull Setup Checklist

1. Pull the latest `main`.
2. Confirm the working tree is clean:

   ```powershell
   git status --short
   ```

   Expected result: no output.
3. Keep the private profile JSON and pricing XLSX outside the repo.
4. If using custom runtime roots, set them before starting the app:

   ```powershell
   $env:QUOTE_DATA_ROOT="<local-runtime-data-root>"
   $env:QUOTE_TMP_ROOT="<local-runtime-temp-root>"
   $env:QUOTE_OUTPUT_ROOT="<local-runtime-output-root>"
   $env:QUOTE_LOG_ROOT="<local-runtime-log-root>"
   $env:KQAG_LOCAL_PRICING_REFERENCES_ROOT="<local-runtime-pricing-root>"
   ```

5. Start the local webapp:

   ```powershell
   python webapp/server.py --host 127.0.0.1 --port 8765
   ```

6. Open:

   ```text
   http://127.0.0.1:8765/
   ```

## Import The Private Quote-Company Profile

1. Open the app locally.
2. Go to the Quote Company step.
3. Use the profile import control.
4. Select `<private-profile-json-outside-repo>`.
5. Confirm the imported profile is selectable as the quote-company profile.
6. Do not paste or report private company fields, bank fields, logo Base64, or
   real file paths in issue notes.

Expected result: the private quote-company profile loads from local/company
runtime storage and can be selected without modifying tracked repo files.

## Import The Private Pricing XLSX

1. Open Pricing Reference Settings.
2. Go to Import.
3. Select `<private-pricing-xlsx-outside-repo>`.
4. Review the imported rows in the app.
5. Save the pricing reference after the preview is valid.
6. Return to Manage and confirm the saved reference is selectable.

Expected result: the private pricing reference is saved in ignored local/runtime
storage and is available to the normal quote flow. The current UI may still use
the temporary wording "local references"; that wording can later become
"workspace pricing references" after login/company platform context exists.

## Generate A Quote With The Normal Flow

1. Start from a fresh quote.
2. Upload the booth render, plan, elevation, fixture schedule, or other
   approved internal test files needed for the quote.
3. Complete Customer and Quote Company details.
4. Confirm the imported private quote-company profile is selected.
5. Confirm the imported private pricing reference is selected.
6. Run AI analysis.
7. Review and confirm the Quote Basis.
8. Open Output.
9. Resolve any pricing review rows.
10. Click Download Excel.

Expected result: quote generation uses the selected company/workspace pricing
reference through the normal quote-facing flow. It should not require an
external temporary generator pack workaround.

## Export XLSX Or PDF

The app's customer-ready master output is `quotation.xlsx`.

For XLSX:

1. Click Download Excel.
2. Open the downloaded workbook in Excel.
3. Confirm Excel opens the workbook without a repair prompt.

For PDF:

1. Open the exported `quotation.xlsx` in Excel.
2. Use Excel's Save as PDF or Print to PDF if an internal PDF copy is needed.
3. Confirm the PDF is generated from the verified XLSX output.

The app does not generate PDF by default in this local test phase.

## Export Verification

After opening the exported workbook, verify:

- The workbook opens without an Excel repair prompt.
- The logo appears if the imported private profile includes a logo.
- GST/company profile values appear correctly.
- Customer-facing wording uses `sqm` for square-metre quantities.
- No internal cost, supplier notes, or hidden private fields are visible unless
  intentionally included by the internal test owner.
- Nonnumeric unit price overrides display safely as `???`.
- Deleting an output row uses the in-app confirmation popup.
- The output does not include an unexpected trailing blank print page.

## Confirm Git Stays Clean

After import, generation, and export, run:

```powershell
git status --short
```

Expected result: no tracked changes caused by private imports or generated
exports. If private/runtime files appear, stop and move them back outside the
repo or into ignored runtime storage before continuing.

Never commit:

- Private profile JSON.
- Private pricing XLSX.
- Exported customer/internal quote workbooks.
- Logo Base64 or image assets from the private profile.
- Private workbook rows/cells.
- Bank details, customer details, or real local usernames/paths.
- Raw screenshots that expose private/customer data unless approved internally.

## Troubleshooting

### Download Fails After Server Restart

If a browser tab was open before the local server restarted, it may hold a stale
session token. The current frontend retries once by refreshing `/api/session`,
but a hard refresh is still the quickest manual recovery.

1. Hard refresh the browser tab.
2. If needed, close and reopen `http://127.0.0.1:8765/`.
3. Recheck that the local server is running.
4. Try Download Excel again.

### Imported Pricing Reference Is Not Selectable

1. Confirm the save step completed in Pricing Reference Settings.
2. Reopen Pricing Reference Settings -> Manage.
3. Confirm the reference appears in the selector.
4. Hard refresh the page if the server was restarted.
5. Confirm `git status --short` did not show private files added under tracked
   repo paths.

### Excel Shows A Repair Prompt

1. Record only the category of the issue, for example "Excel repair prompt on
   export".
2. Do not paste workbook XML, private rows, or private values into GitHub.
3. Attach a sanitized screenshot only if it does not show private/customer data.
4. Keep the broken export outside the repo unless an internal maintainer asks
   for it through an approved private channel.

### AI Analysis Or Import Fails

1. Record the visible reference/error code if shown.
2. Do not paste private upload contents.
3. Note whether the failure happened during analysis, pricing import preview,
   pricing metadata enrichment, quote generation, or Excel download.

## Concise Test Checklist

Use this list for each internal smoke run:

- [ ] Profile import succeeds.
- [ ] Pricing import succeeds.
- [ ] Imported pricing reference is selectable in the normal quote flow.
- [ ] Quote generation succeeds without an external temp generator pack
      workaround.
- [ ] XLSX opens without Excel repair prompt.
- [ ] Logo appears in XLSX export when the private profile includes a logo.
- [ ] GST/company profile values appear correctly.
- [ ] Nonnumeric unit price override displays safely as `???`.
- [ ] Output row delete confirmation popup works.
- [ ] `git status --short` stays clean.

## Reporting Template

```text
Tester name:
Date:
Browser:
App URL:
Branch or commit tested:

Data set used:
Describe without attaching private files or real private paths.

Checklist:
- Profile import succeeds: Pass/Fail
- Pricing import succeeds: Pass/Fail
- Imported pricing reference selectable: Pass/Fail
- Normal quote generation without temp generator pack workaround: Pass/Fail
- XLSX opens without Excel repair prompt: Pass/Fail
- Logo appears when private profile includes logo: Pass/Fail
- GST/company profile values appear correctly: Pass/Fail
- Nonnumeric unit price override displays as ???: Pass/Fail
- Output row delete confirmation popup works: Pass/Fail
- git status --short stays clean: Pass/Fail

Export file opened successfully: Yes/No
Git clean: Yes/No

Screenshot guidance:
Do not include bank/private/customer data unless approved internally.

Issues found:
Describe issue category, visible error/reference code, browser, and step.
Do not paste private company details, bank details, workbook cells, logo Base64,
customer data, or real private file paths.

Follow-up needed:
```
