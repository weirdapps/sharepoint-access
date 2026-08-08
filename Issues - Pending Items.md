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

### P2: Confirm GitHub remote and visibility

The repo is local only. It must be **private**: it documents tenant behaviour
and internal endpoints. Confirm the org (`weirdapps`, matching the siblings) and
create the remote before any push. The PII gauntlet is a required CI job, so the
workflows expect a remote.

### P3: Cold-profile login is unproven

The capture code being ported from `outlook-access` assumes a **warm** context
whose Microsoft SSO cookies were just established by an Outlook sign-in, so it
navigates and expects silent completion. Standalone, the first `login` meets the
full interactive redirect chain including MFA. This is the single largest
unknown in the design and is why phase 1 exists on its own.

### P4: Remove `ExitCode.NotImplemented` once phase 3 lands

Code 7 exists only while the CLI is a scaffold. It diverges from the 0-6 range
shared with `outlook-cli` and `teams-cli` and must not outlive the scaffold.

### P5: Sonar project must be created

`sonar-project.properties` declares `weirdapps_sharepoint-access`. The
SonarCloud project does not exist yet, so `sonarcloud.yml` will fail until it is
created or the workflow is removed.

## Completed

None yet.
