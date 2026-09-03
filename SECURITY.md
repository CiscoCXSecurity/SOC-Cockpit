# Security Policy

## Reporting a Vulnerability

Please report security vulnerabilities responsibly. **Do not** open a public
GitHub issue for security problems.

- For Cisco products and open-source projects, report to the Cisco Product
  Security Incident Response Team (PSIRT): https://www.cisco.com/go/psirt
- Alternatively, use GitHub's private vulnerability reporting on this
  repository (Security tab → "Report a vulnerability").

Please include:

- A description of the vulnerability and its impact.
- Steps to reproduce or a proof of concept.
- Affected version, commit, or configuration.

We will acknowledge your report and work with you on remediation and
coordinated disclosure.

## Supported Versions

Security fixes are applied to the latest released version on the default
branch. Older versions are not guaranteed to receive patches.

## Handling Secrets and Data

- Never commit real credentials. Use `.env` (git-ignored) and copy from
  `.env.example`.
- The application stores operational data in a local SQLite database under
  `data/` and uploaded files under `uploads/`; both are git-ignored and must
  not be committed.

## Known Advisories

The following advisory is reported by `npm audit` and has been assessed as not
exploitable in this project. It is tracked here rather than suppressed.

### GHSA-w5hq-g745-h8pq — `uuid` missing buffer bounds check

- **Severity:** Moderate
- **Path:** `exceljs` -> `uuid@8.3.2`
- **Status:** Accepted, not reachable.

The advisory affects the v3, v5 and v6 generators when a caller supplies a `buf`
argument. `exceljs` imports only the v4 generator and calls it with no arguments
(`lib/xlsx/xform/sheet/cf-rule-ext-xform.js`), so the vulnerable code path is
never entered.

No upstream fix is available: `exceljs@4.4.0` is the current release and it
requires `uuid@^8.3.0`. Forcing a patched major of `uuid` through an `overrides`
entry would place an incompatible API inside that range, and `npm audit fix
--force` resolves this advisory only by downgrading `exceljs` to 3.4.0, a
breaking change to spreadsheet import and export.

This will be revisited when `exceljs` ships a release that depends on
`uuid@>=11.1.1`.
