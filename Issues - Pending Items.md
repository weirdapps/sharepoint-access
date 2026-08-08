# Issues - Pending Items

Pending items first, most critical at the top. Completed items below.

## Pending

### P3: atm-recon must be migrated before anything is removed

`atm-recon/scripts/sharepoint_download.py` reads
`~/.outlook-cli/sharepoint-session.json` **directly**, rather than shelling out
to `outlook-cli`, so it breaks the instant the session file moves. Its
`CLAUDE.md` and `README.md` document the same path and login command.

It also authenticates against the `-my` ODfB host, which is why OneDrive for
Business is in scope (`project-design.md` §2.2) despite an earlier revision
listing it as a non-goal.

Recorded as F-500 and gating the rest of phase 5.

**Process note:** the search that originally concluded "no consumers" omitted
the `sharepoint-session` term and returned a false negative. Search all three
terms (`sharepoint-host`, `download-sharepoint-link`, `sharepoint-session`)
before declaring any repo clean.

### P8: Live write path is unproven

`mkdir` and `put` have complete unit coverage and the write capability is proven
(the auth-check write probe mints a digest on both hosts), but no file has been
created in a live library. Needs an operator-nominated scratch folder, ideally
in the `-my` OneDrive rather than a shared team library.

### P4: Cold-profile login is unproven

The capture code being ported from `outlook-access` assumes a **warm** context
whose Microsoft SSO cookies were just established by an Outlook sign-in, so it
navigates and expects silent completion. Standalone, the first `login` meets the
full interactive redirect chain including MFA. This is the single largest
unknown in the design and is why phase 1 exists on its own.

### P7: Whole-file buffering caps practical file size

`put` reads the entire local file into memory and `getBinary` buffers the whole
response. The chunked uploader slices an already-resident Buffer, so chunking
bounds the request size but not memory: a multi-gigabyte file exhausts the heap.
Streaming both directions is the fix. See `docs/design/security-audit.md` L1.

### P6: Sonar project must be created

`sonar-project.properties` declares `weirdapps_sharepoint-access`. The
SonarCloud project does not exist yet, so `sonarcloud.yml` will fail until it is
created or the workflow is removed.

## Completed

### P1: Config no-fallback exception (resolved 2026-08-08)

Recorded in `CLAUDE.md` before implementation, as the rule requires. The four
plumbing settings have defaults; `host` does not and raises `CONFIG_MISSING`,
because it identifies which tenant is addressed.

### P2: GitHub remote (resolved 2026-08-08)

Created public at `weirdapps/sharepoint-access`, matching `outlook-access` and
`teams-access`. The PII gauntlet is what makes public safe and runs as a
required CI job.

### P5: ExitCode.NotImplemented removed (resolved 2026-08-08)

The scaffold-only code 7 is gone; the enum is back to the 0-6 range shared with
`outlook-cli` and `teams-cli`, with a test asserting it stays that way.
