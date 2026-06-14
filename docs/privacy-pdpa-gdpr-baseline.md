# Privacy, PDPA, and GDPR Baseline

This document is the product privacy baseline for the Swooshz Quote Generator. Production implementation, deployment, database design, AI-provider integration, logging, and user-facing UI must conform to this baseline unless the change is reviewed and this document is updated in the same change.

This is an engineering and product baseline, not legal advice. Before external launch, have counsel review the live deployment, customer contracts, sub-processors, retention periods, and cross-border transfer terms.

## Scope

This baseline covers:

- Uploaded booth render/reference images.
- Customer, project, quotation, company header, and quotation output data.
- Pricing references and profile presets.
- AI analysis prompts, model responses, and generated quote basis/output rows.
- Local and production logs, usage records, error references, and support diagnostics.

## Required Production Behaviors

- Only collect data required to analyze booth images, draft quote basis rows, manage pricing references, and generate quotation files.
- Show users a privacy notice before production use and keep it accessible from the app shell.
- Treat uploaded images, customer details, generated quotations, and pricing references as confidential business data.
- Do not expose internal cost, markup, hidden pricing reference fields, stack traces, secrets, raw prompts, raw provider errors, request headers, or private file paths in customer-facing UI.
- Use role-based access control for pricing reference management, profile management, usage records, and support diagnostics.
- Keep destructive actions reversible where practical, or require explicit confirmation when they cannot be undone.
- Assign support-safe error or ticket references for user-visible failures; keep detailed diagnostics server-side only.
- Keep logs privacy-minimized and redact secrets, tokens, API keys, Authorization headers, cookies, raw images, and long customer text.
- Store usage records per authenticated user/account in production, including provider, model, mode, image count, token usage when available, estimated cost, status, and error reference.
- Define and enforce retention periods for uploads, generated files, AI payloads, usage records, and logs before production launch.
- Provide a user support path for access, correction, deletion, portability, restriction/objection, and consent withdrawal requests where applicable.
- Do not add analytics, trackers, pixels, external scripts, or third-party embeds without explicit privacy review and user notice/consent where required.
- Document every AI provider and infrastructure sub-processor before production launch.
- Review cross-border transfer requirements before sending personal data or confidential business data to any provider outside the user's jurisdiction.

## Current Local Development Behavior

- The local app stores workflow state in the browser, writes generated files under local repo output directories, and writes diagnostics under the root `_logs` folder.
- Uploaded images, uploaded PDFs, rendered PDF page images, and quote data may be sent to the configured AI provider when analysis is run.
- OpenAI is the current full PDF/image draft-analysis provider; DeepSeek may be configured only for text-only quote-basis chat and pricing-reference import normalization.
- Pricing references saved through the local app are stored as repo pricing-reference packs.
- Error UI should show a short retry/support message with an error reference, while detailed provider errors stay in local logs.

## Production Launch Checklist

- Publish the current privacy notice and link it from the app shell.
- Replace local-only role simulation with authenticated users and durable account/user IDs.
- Add database-backed user/account data partitioning with row-level access controls.
- Add usage/cost logging per user and expose an admin usage view.
- Add retention/deletion jobs for uploads, generated quotes, logs, and AI payload records.
- Finalize support contact details and data request workflow.
- Complete AI provider data processing and sub-processor review.
- Review PDPA/GDPR lawful basis, consent/notice language, data transfer clauses, and DPA/SCC needs with counsel.
