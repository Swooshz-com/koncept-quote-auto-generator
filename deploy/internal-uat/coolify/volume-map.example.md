# KQAG Internal UAT Volume Map

Map persistent storage for these container paths in the single Coolify UAT app.
Names below are examples; use the prepared host's approved volume naming
convention.

| Purpose | Env var | Container path | Example Coolify volume |
| --- | --- | --- | --- |
| Runtime company/profile/pricing/session data | `QUOTE_DATA_ROOT` | `/var/lib/kqag/data` | `kqag-uat-data` |
| Generated quote outputs | `QUOTE_OUTPUT_ROOT` | `/var/lib/kqag/output` | `kqag-uat-output` |
| Temporary job/work files | `QUOTE_TMP_ROOT` | `/var/lib/kqag/tmp` | `kqag-uat-tmp` |
| Runtime logs | `QUOTE_LOG_ROOT` | `/var/log/kqag` | `kqag-uat-logs` |

These paths may contain private internal quote workflow data. Do not expose them
as public static directories, do not browse them through the app, and do not
commit their contents to git.

Multi-instance deployment is intentionally out of scope. Durable per-user or
per-account storage partitioning belongs to future platform work, not this
single-instance internal UAT adapter.
