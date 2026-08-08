# Issues - Pending Items

Pending items first, most critical at the top. Completed items below.

## Pending

### P9: Strip SharePoint from outlook-access (F-501), after one clean cycle

Everything this was gated on is done. Held back deliberately: the session
producer changed from `outlook-cli auth-renew --sharepoint-host` to
`sharepoint-cli auth-renew` at 22:30 on 2026-08-08, and both run in parallel
right now. Removing the fallback before the new producer has completed an
unattended cycle is the one remaining move that could cause an outage, and it
would only surface five days later when the rolling window lapsed.

Green-light checks:

1. `~/.sharepoint-cli/session.json` `tokenExpiresAt` has rolled forward on its
   own, with no `sharepoint-cli auth-renew FAILED` in
   `~/Library/Logs/token-sync-vps.log`.
2. The 02:00 `sb-attachments` run on the VPS shows SharePoint fetches
   succeeding rather than `fetched: 0`.
3. `sharepoint-cli auth-check` returns `ok` on both hosts, from both machines.

Then: drop the `--sharepoint-host` half of `sync-tokens-to-vps.sh` line 59,
change `~/.claude/hooks/outlook-reauth.sh:25` to a bare `outlook-cli login`
and delete its `SHAREPOINT_HOST` line, and remove `sharepoint-capture.ts`,
`sharepoint-client.ts`, `sharepoint-schema.ts`, `download-sharepoint-link` and
the `--sharepoint-host` flag from `outlook-access`. Rebuild and reinstall
outlook-cli on **both** machines, or the removal is not actually live.

### P7: Whole-file buffering caps practical file size

`put` reads the entire local file into memory and `getBinary` buffers the whole
response. The chunked uploader slices an already-resident Buffer, so chunking
bounds the request size but not memory: a multi-gigabyte file exhausts the heap.
Streaming both directions is the fix. See `docs/design/security-audit.md` L1.

## Completed

### P3: atm-recon migrated (resolved 2026-08-08)

Ported to `sharepoint-cli` and verified: 189 files downloaded, xlsx validated
as real workbooks. It had in fact been broken beforehand with
`KeyError: 'bearer'`, so this was a repair as well as a port.

### P4: cold-profile login (resolved 2026-08-08, and the premise was wrong)

No interactive MFA was needed. `auth-renew` succeeded headlessly from a cold
profile and the resulting profile holds `ESTSAUTHPERSISTENT` valid to
2026-11-06. `sharepoint-cli` renews itself; it is no longer downstream of
`outlook-cli`.

### P8: live write path (resolved 2026-08-08)

Proven in a scratch folder in the personal OneDrive, then recycled: mkdir, put
with a Greek plus apostrophe filename, ls, get with byte-identical round-trip,
overwrite, and a 12 MB chunked upload also byte-identical.

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

### P6: SonarCloud (resolved 2026-08-08)

Not needed after all. The first push ran `SonarCloud Analysis` to success
alongside `CI`, so the project resolved without manual creation.
