# Current CI/CD Status

Branch: `codex/ci-settings-basis-fixes`

## Configured

- `.github/workflows/ci.yml` runs on pull requests, pushes to `main`, and manual dispatch.
- Secret scanning runs first with Gitleaks.
- Dependency review runs on pull requests.
- Validation runs only after the security gates pass or the pull-request-only dependency review is skipped.
- Validation checks JavaScript syntax, Python syntax, Python unit tests, Playwright AI basis-chat stress smoke, and the main Playwright app smoke test.

## Secrets

- CI does not require OpenAI, Gemini, OIDC, deployment, or production secrets.
- `.env.example` contains blank placeholders only.
- Local `.env` files stay ignored and must not be committed.

## Not Configured

- No deployment workflow is enabled.
- No production environment mutation is performed by CI.
- CodeQL is not enabled in this first pass to avoid requiring repository code-scanning features before the base CI gate is stable.

## Next Manual Steps

- Push this branch and open a pull request when ready.
- After the first green pull request, enable required branch protection for `CI / Secret scan`, `CI / Dependency review`, and `CI / Validate app`.
