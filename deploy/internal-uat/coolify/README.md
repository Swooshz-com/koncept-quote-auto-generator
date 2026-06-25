# KQAG Internal UAT Coolify Adapter

This folder contains KQAG-specific templates for an already-prepared Coolify
host. It intentionally does not describe generic VPS setup, Coolify
installation, SSH, firewall, DNS, TLS, or server maintenance.

Use with:

- Docs: `docs/internal-uat-coolify-deploy.md`
- Env template: `deploy/internal-uat/coolify/kqag.uat.env.example`
- Volume map: `deploy/internal-uat/coolify/volume-map.example.md`

Recommended Coolify app settings:

- Runtime/buildpack: Python using `requirements.txt`
- Start command: `python webapp/server.py`
- Port: `8765`, or the value supplied by `PORT`
- Healthcheck path: `/api/health`
- Instance count: `1`

Before starting deploy-mode UAT, set the env values in Coolify secrets or
environment management and run:

```powershell
python webapp\server.py --check-deploy-uat-env
```

Keep populated env files, real OIDC values, private profile/pricing files,
runtime data, and generated quote exports out of git.
