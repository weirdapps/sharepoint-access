# Issues - Pending Items

Pending items first, most critical at the top. Completed items below.

## Pending

### P1: Decide the config no-fallback exception before phase 1

`CLAUDE.md` carries the global rule that a missing configuration setting must
raise rather than fall back to a default. `outlook-access` recorded an explicit
exception for three runtime-plumbing settings (`httpTimeoutMs`,
`loginTimeoutMs`, `chromeChannel`) because forcing them on every invocation
traded ergonomics for safety the rule was not protecting.

This repo will need the same three, plus a tenant `host`. Decide and record
before writing `src/config/load.ts`:

- Do the three plumbing settings inherit the same exception? Probably yes, for
  consistency across the sibling CLIs.
- Is `host` plumbing or identity? It is arguably **identity**, so it should have
  no default and should raise when absent, even though the tenant only has one
  host in practice. A default here would silently point at the wrong tenant.

Per the rule, any exception must be written into `CLAUDE.md` **before** it is
implemented.

### P2: Create the GitHub remote

The repo is local only. It should be **public** in `weirdapps`, matching
`outlook-access` and `teams-access`, which are both public. The PII gauntlet is
what makes that safe and runs as a required CI job.

Public is only safe while the docs stay generic. This repo describes the
tenant's security posture (no Bearer issued, `/_api/v2.0` blocked, MCAS in play)
in more detail than the siblings do, so it must never also name the
organisation. The gauntlet blocks the tenant host but does **not** block the
bank's name, and `outlook-access` already carries that name in tracked files.
Keep writing `<tenant>.sharepoint.com` and avoid the org name here.

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

### P4: Cold-profile login is unproven

The capture code being ported from `outlook-access` assumes a **warm** context
whose Microsoft SSO cookies were just established by an Outlook sign-in, so it
navigates and expects silent completion. Standalone, the first `login` meets the
full interactive redirect chain including MFA. This is the single largest
unknown in the design and is why phase 1 exists on its own.

### P5: Remove `ExitCode.NotImplemented` once phase 3 lands

Code 7 exists only while the CLI is a scaffold. It diverges from the 0-6 range
shared with `outlook-cli` and `teams-cli` and must not outlive the scaffold.

### P6: Sonar project must be created

`sonar-project.properties` declares `weirdapps_sharepoint-access`. The
SonarCloud project does not exist yet, so `sonarcloud.yml` will fail until it is
created or the workflow is removed.

## Completed

None yet.
