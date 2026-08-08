# Project Functions

Functional register for `sharepoint-cli`. Every command and capability is
recorded here as it is designed, with its delivery phase and current state.
Phases are defined in `project-design.md` §11.

**State legend:** `planned` designed but not built. `built` implemented and unit
tested. `verified` additionally proven against the live tenant.

## Phase 1: Authentication

| ID    | Function       | State    | Description                                                                                                                                                                                                  |
| ----- | -------------- | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| F-101 | `login`        | built    | Launch persistent Chrome on `~/.sharepoint-cli/playwright-profile`, navigate to the tenant, wait for interactive sign-in including MFA, capture `FedAuth`/`rtFa` cookies and any Bearer, persist the session |
| F-102 | `auth-renew`   | built    | Headless relaunch on the same profile, silent SSO via `ESTSAUTHPERSISTENT`, re-capture and persist. Short timeout                                                                                            |
| F-103 | `auth-check`   | verified | Probe all three dependencies (read, write, search) and report per-surface status. Non-zero exit if any fails                                                                                                 |
| F-104 | `health-check` | verified | Same three probes with per-probe timings, formatted for cron                                                                                                                                                 |
| F-105 | Session store  | built    | Load and save `~/.sharepoint-cli/session.json`, schema v1: host, bearer?, cookies, capturedAt, tokenExpiresAt                                                                                                |
| F-106 | Browser lock   | built    | `~/.sharepoint-cli/.browser.lock` preventing concurrent Playwright launches against one profile                                                                                                              |

**F-103 note.** Probing a single surface is the defect being fixed in
`teams-access` (`project-design.md` appendix A). All three are probed here by
design: `GET /_api/web` for reads, `POST /_api/contextinfo` for writes,
`GET /_api/search/query` for search.

## Phase 2: Read surface

| ID    | Function           | State    | Description                                                                                                                                                            |
| ----- | ------------------ | -------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| F-201 | `ls <path>`        | verified | List a folder's subfolders and files with size and modified time                                                                                                       |
| F-202 | `get <path\|url>`  | built    | Download a file by server-relative path or absolute URL. Replaces `outlook-cli download-sharepoint-link`                                                               |
| F-203 | `search <query>`   | verified | Search files and sites. Site discovery is `search "contentclass:STS_Site"`, there is no `sites` command                                                                |
| F-204 | `libraries [site]` | verified | List non-hidden document libraries (`BaseTemplate eq 101`) in a site                                                                                                   |
| F-205 | Path handling      | built    | OData apostrophe doubling, UTF-8 percent-encoding, encode-layer separation, traversal rejection, parent/leaf split                                                     |
| F-206 | Dual-host support  | verified | `--host` accepts the team-sites host or the `-my` ODfB host. One session covers both, so no separate code path, but both must be smoke tested. Required by `atm-recon` |

## Phase 3: Write surface

| ID    | Function               | State    | Description                                                                                                     |
| ----- | ---------------------- | -------- | --------------------------------------------------------------------------------------------------------------- |
| F-301 | Request digest         | verified | Fetch and cache `X-RequestDigest`, expiring at `FormDigestTimeoutSeconds` minus 60s. Never persisted            |
| F-302 | `mkdir <path>`         | built    | Create a folder. Parents must exist, no implicit `-p`                                                           |
| F-303 | `put <local> <folder>` | built    | Upload, routed by size: single request below 10 MB, chunked `StartUpload`/`ContinueUpload`/`FinishUpload` above |
| F-304 | Upload cleanup         | built    | Delete the partial target on any chunk failure so retries start clean                                           |

## Phase 4: MCP exposure

| ID    | Function       | State   | Description                                                                                              |
| ----- | -------------- | ------- | -------------------------------------------------------------------------------------------------------- |
| F-401 | `files` plugin | planned | New `plessas-marketplace/plugins/files/`, thin subprocess bridge over the CLI, no independent auth logic |

Tools exposed: `sharepoint_list`, `sharepoint_get`, `sharepoint_put`,
`sharepoint_mkdir`, `sharepoint_search`, `sharepoint_libraries`.

## Phase 5: Removal from outlook-access

| ID    | Function            | State   | Description                                                                                                                                                                      |
| ----- | ------------------- | ------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| F-500 | Migrate `atm-recon` | planned | Port `scripts/sharepoint_download.py` to the new session path, update its `CLAUDE.md` and `README.md`, verify it still pulls its three folders. **Gate for the rest of phase 5** |
| F-501 | Strip SharePoint    | planned | Remove `--sharepoint-host`, `download-sharepoint-link`, `sharepoint-capture.ts`, `sharepoint-client.ts`, `sharepoint-schema.ts` and their tests and docs                         |
| F-502 | Update re-auth docs | planned | Global `CLAUDE.md` and `plessas-trading-stack` newsfeed `SKILL.md` re-auth lines point at `sharepoint-cli login`                                                                 |
| F-503 | Session migration   | planned | `~/.outlook-cli/sharepoint-session.json` to `~/.sharepoint-cli/session.json`, or simply re-run `login`                                                                           |

Gated on phases 1 to 3 being `verified`, and F-500 gates F-501 to F-503.

## Cross-cutting

| ID    | Function      | State | Description                                                                                            |
| ----- | ------------- | ----- | ------------------------------------------------------------------------------------------------------ |
| F-001 | Exit codes    | built | `src/util/exit-codes.ts`, mirroring `teams-access` for 0-6, plus temporary 7 `NotImplemented`          |
| F-002 | CLI skeleton  | built | `src/cli.ts`, registers every planned command so `--help` is accurate and unbuilt commands fail loudly |
| F-003 | Error mapping | built | 401 auth-required, 403 digest-stale versus access-denied, 404, 423 locked, 507 quota                   |

## Audit and integration

| ID    | Function                  | State | Description                                                                                                                  |
| ----- | ------------------------- | ----- | ---------------------------------------------------------------------------------------------------------------------------- |
| F-901 | Prototype-pollution guard | built | `search` builds results on a null prototype and skips `__proto__`/`constructor`/`prototype`. See `security-audit.md` AUDIT-1 |
| F-902 | Relative-path host guard  | built | `client.url()` requires a single leading `/`, so a path cannot concatenate into another host. AUDIT-2                        |
| F-903 | Digest recursion guard    | built | `contextInfo()` bypasses the digest-carrying POST path; the cycle is unrepresentable in the type                             |
| F-904 | Wired integration tests   | built | `integration.spec.ts` exercises the real client plus real digest cache, mocking only `fetch`                                 |

## Verification note

`verified` above means proven against the live tenant on 2026-08-08, on both
the team-sites and `-my` hosts where applicable.

`login` and `auth-renew` are `built`, not `verified`: interactive sign-in needs
a human at an MFA prompt, and headless renewal needs a profile that only a
completed interactive login can create. Phases 2 and 3 were verified using the
session seeded from `outlook-cli`.

`get`, `mkdir` and `put` are `built`: their unit coverage is complete and the
write **capability** is proven (the `auth-check` write probe mints a digest
successfully on both hosts), but no file has been created in a live library.
That needs an operator-nominated scratch folder.
