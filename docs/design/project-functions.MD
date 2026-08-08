# Project Functions

Functional register for `sharepoint-cli`. Every command and capability is
recorded here as it is designed, with its delivery phase and current state.
Phases are defined in `project-design.md` §11.

**State legend:** `planned` designed but not built. `built` implemented and unit
tested. `verified` additionally proven against the live tenant.

## Phase 1: Authentication

| ID    | Function       | State   | Description                                                                                                                                                                                                  |
| ----- | -------------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| F-101 | `login`        | planned | Launch persistent Chrome on `~/.sharepoint-cli/playwright-profile`, navigate to the tenant, wait for interactive sign-in including MFA, capture `FedAuth`/`rtFa` cookies and any Bearer, persist the session |
| F-102 | `auth-renew`   | planned | Headless relaunch on the same profile, silent SSO via `ESTSAUTHPERSISTENT`, re-capture and persist. Short timeout                                                                                            |
| F-103 | `auth-check`   | planned | Probe all three dependencies (read, write, search) and report per-surface status. Non-zero exit if any fails                                                                                                 |
| F-104 | `health-check` | planned | Same three probes with per-probe timings, formatted for cron                                                                                                                                                 |
| F-105 | Session store  | planned | Load and save `~/.sharepoint-cli/session.json`, schema v1: host, bearer?, cookies, capturedAt, tokenExpiresAt                                                                                                |
| F-106 | Browser lock   | planned | `~/.sharepoint-cli/.browser.lock` preventing concurrent Playwright launches against one profile                                                                                                              |

**F-103 note.** Probing a single surface is the defect being fixed in
`teams-access` (`project-design.md` appendix A). All three are probed here by
design: `GET /_api/web` for reads, `POST /_api/contextinfo` for writes,
`GET /_api/search/query` for search.

## Phase 2: Read surface

| ID    | Function           | State   | Description                                                                                                                                                            |
| ----- | ------------------ | ------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| F-201 | `ls <path>`        | planned | List a folder's subfolders and files with size and modified time                                                                                                       |
| F-202 | `get <path\|url>`  | planned | Download a file by server-relative path or absolute URL. Replaces `outlook-cli download-sharepoint-link`                                                               |
| F-203 | `search <query>`   | planned | Search files and sites. Site discovery is `search "contentclass:STS_Site"`, there is no `sites` command                                                                |
| F-204 | `libraries [site]` | planned | List non-hidden document libraries (`BaseTemplate eq 101`) in a site                                                                                                   |
| F-205 | Path handling      | planned | OData apostrophe doubling, UTF-8 percent-encoding, encode-layer separation, traversal rejection, parent/leaf split                                                     |
| F-206 | Dual-host support  | planned | `--host` accepts the team-sites host or the `-my` ODfB host. One session covers both, so no separate code path, but both must be smoke tested. Required by `atm-recon` |

## Phase 3: Write surface

| ID    | Function               | State   | Description                                                                                                     |
| ----- | ---------------------- | ------- | --------------------------------------------------------------------------------------------------------------- |
| F-301 | Request digest         | planned | Fetch and cache `X-RequestDigest`, expiring at `FormDigestTimeoutSeconds` minus 60s. Never persisted            |
| F-302 | `mkdir <path>`         | planned | Create a folder. Parents must exist, no implicit `-p`                                                           |
| F-303 | `put <local> <folder>` | planned | Upload, routed by size: single request below 10 MB, chunked `StartUpload`/`ContinueUpload`/`FinishUpload` above |
| F-304 | Upload cleanup         | planned | Delete the partial target on any chunk failure so retries start clean                                           |

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

| ID    | Function      | State   | Description                                                                                            |
| ----- | ------------- | ------- | ------------------------------------------------------------------------------------------------------ |
| F-001 | Exit codes    | built   | `src/util/exit-codes.ts`, mirroring `teams-access` for 0-6, plus temporary 7 `NotImplemented`          |
| F-002 | CLI skeleton  | built   | `src/cli.ts`, registers every planned command so `--help` is accurate and unbuilt commands fail loudly |
| F-003 | Error mapping | planned | 401 auth-required, 403 digest-stale versus access-denied, 404, 423 locked, 507 quota                   |
