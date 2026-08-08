# Security Audit: 2026-08-08

Audit of `sharepoint-cli` at commit `ead234b`, covering phases 1 to 3. Scope:
credential handling, request targeting, injection, path traversal, filesystem
permissions, error-message leakage, and dependencies.

Two findings were fixed during the audit. Both have regression tests in
`test_scripts/security.spec.ts` that fail if the defence is removed.

## Threat model

This CLI holds a live SharePoint session for a corporate tenant, in cookie form
with no expiry shorter than several days. The assets worth protecting, in order:

1. **The session cookies.** They grant the operator's full document-store
   access. Exfiltrating them is worse than any single file.
2. **Write integrity.** The CLI can create and overwrite documents in live
   libraries.
3. **The operator's local filesystem**, via download paths.

The realistic adversary is not a network attacker (everything is TLS to
Microsoft) but **attacker-influenced data**: a document name, a share URL in an
email, or a search result crafted by anyone who can put a file in a library the
operator can read.

## Findings fixed

### AUDIT-1: Prototype pollution via search result keys (medium)

`runSearch` built its per-row object by assigning server-supplied cell keys:

```typescript
const out: Record<string, string> = {};
out[cell.Key] = cell.Value; // cell.Key is whatever SharePoint returns
```

A search result row keyed `__proto__`, `constructor`, or `prototype` reaches
`Object.prototype`. This is reachable rather than theoretical: managed
properties in the search index are influenced by document metadata, and anyone
who can add a document to an indexed library can influence them.

Fixed by building on a null prototype and skipping the three dangerous keys,
then spreading back to a plain object so JSON serialisation is unchanged.

### AUDIT-2: Relative path could retarget the request host (low, unreachable)

`SharepointClient.url()` concatenated any non-absolute path onto the host:

```typescript
return `https://${this.session.host}${pathOrUrl}`;
```

A path lacking a leading slash, such as `evil.com/x`, produces
`https://x.sharepoint.comevil.com/x`, a **different host** that would still
receive the `Cookie` header. A protocol-relative `//evil.com/x` is worse.

No current caller can reach this: every path is built by `folderApi`, `fileApi`,
or a literal `/_api/...` string. It was fixed anyway, because the invariant is
invisible at the call sites and one future caller passing user input would
reintroduce a credential-exfiltration bug.

## Verified, no change needed

| Area                                      | Finding                                                                                                                                                                                                                                                                                                                                                                               |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Cookie exfiltration via absolute URLs** | `assertSharepointUrl` runs **before** any header is attached. Requires `https:` and a hostname that is exactly `sharepoint.com` or ends in `.sharepoint.com`, so `evilsharepoint.com` and `x.sharepoint.com.evil.com` are both refused. This matters because `get` accepts share URLs that arrive from email.                                                                         |
| **Tenant host validation**                | `HOST_RE` is anchored at both ends and permits only `<label>.sharepoint.com`, rejecting embedded credentials (`evil.com@x…`), ports, and paths.                                                                                                                                                                                                                                       |
| **Path traversal**                        | `normalizeServerRelative` rejects `..` **after** collapsing duplicate slashes, so `/a//../b` cannot smuggle a segment past the scan, and rejects protocol-relative input **before** collapsing, so `//host/x` cannot be disguised as `/host/x`. Upload leaf names go through `encodeLeaf`, which rejects `/`, `.` and `..`, so `put ../../evil.docx` cannot escape the target folder. |
| **OData injection**                       | Every value interpolated into an OData string literal is escaped by doubling apostrophes, then percent-encoded, in that order. Applies to paths, upload leaf names, and the free-text search query. Verified against the live tenant that both the bare `''` and a percent-encoded form resolve to the same single-apostrophe filename.                                               |
| **Control-character injection**           | Paths and leaf names reject C0 and C1 control characters, so CR/LF cannot be smuggled into a request line.                                                                                                                                                                                                                                                                            |
| **Credential leakage in errors**          | `SharepointHttpError`'s message carries status and URL only; the response body stays on a property and is never printed by the output layer. Regression-tested that neither the bearer nor the cookie value appears in `message` or `stack`.                                                                                                                                          |
| **Credentials on disk**                   | Session written 0600 inside a 0700 directory, via temp-then-rename so a crash cannot leave a truncated file. The lock file is 0600 via `O_CREAT\|O_EXCL`. The Playwright profile directory is 0700.                                                                                                                                                                                   |
| **Digest handling**                       | The CSRF digest is process-local and never persisted. Its lifetime (1800s) is shorter than most cron gaps, so writing it to disk would add attack surface for nothing.                                                                                                                                                                                                                |
| **Randomness**                            | Upload session IDs use `randomUUID` from `node:crypto`, not `Math.random`.                                                                                                                                                                                                                                                                                                            |
| **Lock safety**                           | Advisory lock uses `O_CREAT\|O_EXCL` and reclaims only when the recorded PID is provably dead (`ESRCH`); `EPERM` is treated as alive so another user's process is never clobbered.                                                                                                                                                                                                    |
| **Secrets in the repo**                   | `gitleaks` and a PII gauntlet run as pre-commit hooks **and** as required CI jobs. The gauntlet blocks the tenant hostname, which caught a real leak in the design doc during this work.                                                                                                                                                                                              |
| **Dependencies**                          | Two runtime dependencies: `commander` and `playwright` (optional, lazy-loaded so read commands never initialise it). `npm audit --omit=dev` reports 0 vulnerabilities. 125 packages installed in total, all dev tooling otherwise.                                                                                                                                                    |
| **Redirects**                             | `redirect: 'follow'` is used, but the fetch specification strips `Cookie` and `Authorization` on cross-origin redirects, and undici implements this. A same-origin redirect within `*.sharepoint.com` retaining cookies is intended behaviour.                                                                                                                                        |

## Accepted limitations

These are real and documented rather than fixed, because fixing them is a
design change rather than a patch. Tracked in `Issues - Pending Items.md`.

### L1: Whole-file buffering (availability, not confidentiality)

`put` reads the entire local file into memory with `readFile`, and `getBinary`
buffers the whole response with `arrayBuffer()`. The chunked uploader slices an
already-resident `Buffer`, so chunking bounds the **request** size but not
memory. A multi-gigabyte file will exhaust the heap.

Not a security hole on a single-operator CLI, but it caps the practical file
size well below what SharePoint accepts. Streaming both directions is the fix
and belongs in a later phase.

### L2: JWT signature is not verified

`deriveTokenExpiry` decodes the bearer payload without verifying the signature.
This is acceptable: the token came from the operator's own browser session, it
is used only to compute a display expiry, and no authorisation decision depends
on it. Malformed input falls through to the cookie-derived window rather than
throwing.

### L3: `get --out` follows symlinks

The output path is resolved and written without checking for an existing
symlink, so writing to a path an attacker has pre-symlinked would write through
it. The path is supplied by the operator on their own machine, so this is the
ordinary behaviour of every download tool.

### L4: Search returns more fields than requested

`selectproperties` adds to SharePoint's default property set rather than
restricting it, so results carry substantial index metadata (`DocId`, `SiteId`,
`WebId`, rank information). Not a vulnerability, but the output is verbose and
callers should not assume it is limited to the requested properties.

## Note on test coverage

The digest recursion bug found during live testing (`postJson` asked the digest
provider for a header, whose only source was `postJson`) survived 214 passing
unit tests because **every test mocked at the client boundary**. The object
graph the CLI actually builds was never exercised.

`test_scripts/integration.spec.ts` now wires the real client to the real digest
cache and mocks only `fetch`. Any future component that participates in the
request path should get a test at that layer, not only at its own boundary.

---

## Round 2: 2026-08-08, after live and cross-machine testing

The first pass audited code. This pass audited the running system, and found
things static review could not.

### Fixed

**Sub-site scoping (correctness, high).** Every `_api` call ran against the
host root web. SharePoint answers `200` with EMPTY collections for a path the
queried web does not own, and for a path that does not exist at all, so the CLI
reported real folders as empty and typos as empty rather than erroring.
Everything under `/sites/*` and `/personal/*` was silently unreachable. Fixed by
deriving the web from the path, requesting `Exists`, and adding `--site`.

**Web-scoped digests (correctness, high).** `X-RequestDigest` is per-web. A
root-minted digest is rejected `403` by a sub-site _carrying the stale-digest
marker_, so it presented as expiry and the retry could never succeed. Every
write outside the root web was broken. `DigestCache` is now keyed by web.

**Prototype pollution via search keys** and **relative paths able to retarget
the request host** were fixed in round 1 and remain covered by
`security.spec.ts`.

### Availability findings

**undici connect timeout.** Fixed at 10s without a dispatcher. Measured
connects to some front-ends on this tenant range from 1.5s to 42s, so roughly
two requests in three failed. Fixed with an Agent, a 60s overall timeout, and a
bounded retry for connection-level failures only.

**Cross-version dispatcher incompatibility.** Node 24 on the VPS rejects a
standalone undici Agent that Node 26 on the Mac accepts, failing every request
instantly. The client now drops the dispatcher and retries. Caught only by
running `auth-check` on the VPS; the Mac was green throughout.

**A race in that fix.** Concurrent requests share one client, so the first to
fail cleared the dispatcher and the rest then skipped their own retry.
`auth-check` runs three probes in parallel and reproduced it exactly. Each
attempt now records whether it used a dispatcher.

### Operational findings, from an adversarial audit of 60 automations

**A transitional mirror that would have destroyed the thing it was preparing
for.** The session mirror was unconditional, so the moment `sharepoint-cli
login` created a real session, the 15-minute job would have overwritten it with
the older `outlook-cli` copy, forever. Now guarded on the profile's existence,
so the block retires itself.

**Silent push failure.** `scp` does not create a missing parent, and every
other `scp` in that script is `2>/dev/null` with no return-code check, so a
push to a non-existent remote directory would have failed silently every 15
minutes. The directory is now ensured and failures are logged.

**Search coverage.** BSD `grep -r` silently skips symlinked files, and
`~/Documents`, `~/Operational` and much of `~/.local/bin` are symlinks. That is
how an earlier hand search missed the live token-sync script. Use
`find -L <dir> -type f -print0 | xargs -0 grep`.
