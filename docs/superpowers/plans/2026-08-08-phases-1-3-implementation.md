# sharepoint-cli Phases 1-3 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `sharepoint-cli` from scaffold to a working CLI that authenticates against SharePoint Online, browses and downloads from document libraries, and uploads files, on both the team-sites host and the OneDrive for Business host.

**Architecture:** Cookie-authenticated classic SharePoint REST (`/_api/web`, `/_api/search`). A Playwright persistent profile captures `FedAuth`/`rtFa` cookies once interactively, then renews them headlessly via `ESTSAUTHPERSISTENT`. Writes carry an `X-RequestDigest` minted from `/_api/contextinfo` and cached in-process. No Microsoft Graph anywhere: the tenant issues no Bearer for SharePoint and returns 403 on `/_api/v2.0`.

**Tech Stack:** TypeScript 6 (CommonJS), Node >= 20, `commander`, `playwright` (optional dep, lazy-loaded), `vitest`.

## Global Constraints

Copied verbatim from `docs/design/project-design.md`. Every task inherits these.

- **Never use Microsoft Graph.** `graph.microsoft.com` needs a Graph-audience Bearer; the tenant issues none and cookies do not authorise it. `/_api/v2.0` returns 403.
- **Use `*ByServerRelativePath(decodedUrl='…')`, never `*ByServerRelativeUrl('…')`.** The `Url` variants mishandle `#` and `%`. Writes use `addUsingPath(DecodedUrl='…')`, never `add(url='…')`.
- **A stale digest returns 403, not 401.** Discriminate on the body containing `security validation for this page is invalid` before reporting a permissions error.
- **A missing Bearer is not an error.** Cookie-auth tenants emit none; `FedAuth`/`rtFa` alone authorise.
- **One session covers both hosts.** Cookies are collected for the parent `sharepoint.com` domain, so `<tenant>.sharepoint.com` and `<tenant>-my.sharepoint.com` both work.
- **Exit codes 0-6 mirror `teams-access` exactly.** 4 = re-authenticate, 5 = upstream. Code 7 (`NotImplemented`) is removed when phase 3 completes.
- **Test files use `.spec.ts`** under `test_scripts/`, enforced by `vitest.config.ts`.
- **Never write the real tenant host into a tracked file.** Use `<tenant>.sharepoint.com`. The PII gauntlet blocks it and runs as a required CI job. Do not name the organisation either.
- **Lazy-load `playwright`** inside functions via `await import('playwright')`, so read-only commands never trigger its module init.
- Repo conventions in `CLAUDE.md` apply: register every function in `docs/design/project-functions.MD`, log defects in `Issues - Pending Items.md`, TypeScript only.

## Development session shortcut

Phase 1's interactive `login` needs a human at an MFA prompt and cannot be verified unattended. A valid session already exists at `~/.outlook-cli/sharepoint-session.json` in exactly the schema this repo uses. Copy it to `~/.sharepoint-cli/session.json` to develop and verify phases 2-3 against the live tenant immediately:

```bash
mkdir -p ~/.sharepoint-cli && chmod 700 ~/.sharepoint-cli
cp ~/.outlook-cli/sharepoint-session.json ~/.sharepoint-cli/session.json
chmod 600 ~/.sharepoint-cli/session.json
```

Interactive `login` and headless `auth-renew` remain unverified until a human runs them. That is called out in the final report, not silently glossed.

## File structure

| File                       | Responsibility                                                        |
| -------------------------- | --------------------------------------------------------------------- |
| `src/config/errors.ts`     | `CliError` base, typed error codes, `ExitWithCode` mapping            |
| `src/config/load.ts`       | Resolve host, timeouts, chrome channel, paths. Flag > env > default   |
| `src/session/schema.ts`    | `SharepointSession` type, parse, serialize                            |
| `src/session/store.ts`     | Load/save `~/.sharepoint-cli/session.json` atomically at 0600         |
| `src/auth/lock.ts`         | Advisory PID lock, ported verbatim from `outlook-access`              |
| `src/auth/capture.ts`      | Playwright capture, interactive and headless                          |
| `src/sharepoint/paths.ts`  | Server-relative path encoding, the two encoding layers                |
| `src/sharepoint/digest.ts` | `X-RequestDigest` fetch and in-process cache                          |
| `src/sharepoint/upload.ts` | Size-routed upload, single-shot and chunked                           |
| `src/http/errors.ts`       | `SharepointHttpError`, status-to-code mapping, digest discrimination  |
| `src/http/client.ts`       | `getJson`, `getBinary`, `postJson`, `postBinary`, retry-once envelope |
| `src/http/types.ts`        | REST response shapes                                                  |
| `src/commands/*.ts`        | One file per command, pure functions returning result objects         |
| `src/output/json.ts`       | Stable JSON stdout, errors to stderr                                  |
| `src/cli.ts`               | Commander wiring only                                                 |

---

### Task 1: Record the config exception, then build config and session

Per `CLAUDE.md`, any exception to the no-fallback config rule must be written down **before** it is implemented. Resolves P1.

**Files:**

- Modify: `CLAUDE.md` (add exception section)
- Create: `src/config/errors.ts`, `src/config/load.ts`, `src/session/schema.ts`, `src/session/store.ts`
- Test: `test_scripts/config.spec.ts`, `test_scripts/session.spec.ts`

**Interfaces:**

- Produces: `CliError`, `ErrorCode`; `loadConfig(overrides): CliConfig` where `CliConfig = { host: string; httpTimeoutMs: number; loginTimeoutMs: number; renewTimeoutMs: number; chromeChannel: string; sessionPath: string; profileDir: string; lockPath: string }`; `SharepointSession`, `parseSession(json): SharepointSession`, `serializeSession(s): string`, `loadSession(path): Promise<SharepointSession | null>`, `saveSession(path, s): Promise<void>`, `defaultSessionPath(): string`

- [ ] **Step 1: Write the exception into `CLAUDE.md`**

Append after the `</structure-and-conventions>` block:

```markdown
## Project-specific exceptions to global rules

### Exception: defaults allowed for four runtime-plumbing settings

`httpTimeoutMs` (30000), `loginTimeoutMs` (300000), `renewTimeoutMs` (30000)
and `chromeChannel` ("chrome") have defaults, mirroring the identical
exception recorded in `outlook-access`. They are operational plumbing, not
secrets or environment-distinguishing identities, so requiring them on every
invocation trades ergonomics for safety the rule does not protect.

`host` is explicitly **not** covered. It identifies which tenant is being
addressed, so it has no default and `loadConfig` raises `CONFIG_MISSING` when
it is absent. A default here would silently point at the wrong tenant.

Precedence: CLI flag > env var (`SHAREPOINT_CLI_*`) > default.
```

- [ ] **Step 2: Write the failing tests**

`test_scripts/config.spec.ts`:

```typescript
import { describe, expect, it } from 'vitest';
import { loadConfig } from '../src/config/load';
import { CliError } from '../src/config/errors';

describe('loadConfig', () => {
  it('raises CONFIG_MISSING when host is absent everywhere', () => {
    expect(() => loadConfig({}, {})).toThrowError(CliError);
    try {
      loadConfig({}, {});
    } catch (e) {
      expect((e as CliError).code).toBe('CONFIG_MISSING');
    }
  });

  it('takes host from the flag ahead of the env var', () => {
    const c = loadConfig({ host: 'a.sharepoint.com' }, { SHAREPOINT_CLI_HOST: 'b.sharepoint.com' });
    expect(c.host).toBe('a.sharepoint.com');
  });

  it('falls back to the env var for host', () => {
    const c = loadConfig({}, { SHAREPOINT_CLI_HOST: 'b.sharepoint.com' });
    expect(c.host).toBe('b.sharepoint.com');
  });

  it('rejects a host that is not a sharepoint.com domain', () => {
    expect(() => loadConfig({ host: 'evil.example.com' }, {})).toThrowError(/sharepoint\.com/);
  });

  it('accepts the -my OneDrive for Business host', () => {
    const c = loadConfig({ host: 'x-my.sharepoint.com' }, {});
    expect(c.host).toBe('x-my.sharepoint.com');
  });

  it('applies documented defaults for the four plumbing settings', () => {
    const c = loadConfig({ host: 'x.sharepoint.com' }, {});
    expect(c.httpTimeoutMs).toBe(30000);
    expect(c.loginTimeoutMs).toBe(300000);
    expect(c.renewTimeoutMs).toBe(30000);
    expect(c.chromeChannel).toBe('chrome');
  });

  it('lets env vars override plumbing defaults', () => {
    const c = loadConfig({ host: 'x.sharepoint.com' }, { SHAREPOINT_CLI_HTTP_TIMEOUT_MS: '1234' });
    expect(c.httpTimeoutMs).toBe(1234);
  });

  it('rejects a non-numeric timeout rather than silently defaulting', () => {
    expect(() =>
      loadConfig({ host: 'x.sharepoint.com' }, { SHAREPOINT_CLI_HTTP_TIMEOUT_MS: 'abc' }),
    ).toThrowError(/SHAREPOINT_CLI_HTTP_TIMEOUT_MS/);
  });
});
```

`test_scripts/session.spec.ts`:

```typescript
import { describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { parseSession, serializeSession, loadSession, saveSession } from '../src/session/store';

const valid = {
  version: 1 as const,
  host: 'x.sharepoint.com',
  cookies: 'FedAuth=aaa; rtFa=bbb',
  capturedAt: '2026-08-08T00:00:00.000Z',
  tokenExpiresAt: '2026-08-13T00:00:00.000Z',
};

describe('parseSession', () => {
  it('accepts a cookie-only session with no bearer', () => {
    expect(parseSession(JSON.stringify(valid)).bearer).toBeUndefined();
  });

  it('accepts a session with a bearer', () => {
    expect(parseSession(JSON.stringify({ ...valid, bearer: 'jwt' })).bearer).toBe('jwt');
  });

  it('rejects an unsupported version', () => {
    expect(() => parseSession(JSON.stringify({ ...valid, version: 2 }))).toThrowError(/version/);
  });

  it('rejects a missing cookies field', () => {
    const { cookies: _drop, ...rest } = valid;
    expect(() => parseSession(JSON.stringify(rest))).toThrowError(/cookies/);
  });

  it('rejects a non-string bearer', () => {
    expect(() => parseSession(JSON.stringify({ ...valid, bearer: 42 }))).toThrowError(/bearer/);
  });

  it('rejects malformed JSON', () => {
    expect(() => parseSession('{oops')).toThrowError(/JSON/);
  });
});

describe('saveSession / loadSession', () => {
  it('round-trips and writes the file 0600 in a 0700 directory', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    const p = path.join(dir, 'nested', 'session.json');
    await saveSession(p, valid);
    expect(await loadSession(p)).toEqual(valid);
    expect(fs.statSync(p).mode & 0o777).toBe(0o600);
    expect(fs.statSync(path.dirname(p)).mode & 0o777).toBe(0o700);
  });

  it('returns null when the file does not exist', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    expect(await loadSession(path.join(dir, 'absent.json'))).toBeNull();
  });

  it('does not leave a .tmp file behind', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    const p = path.join(dir, 'session.json');
    await saveSession(p, valid);
    expect(fs.existsSync(p + '.tmp')).toBe(false);
  });
});

describe('serializeSession', () => {
  it('produces parseable output', () => {
    expect(parseSession(serializeSession(valid))).toEqual(valid);
  });
});
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `npm test`
Expected: FAIL, cannot resolve `../src/config/load` and `../src/session/store`.

- [ ] **Step 4: Implement `src/config/errors.ts`**

```typescript
// src/config/errors.ts
import { ExitCode, ExitWithCode, type ExitCodeValue } from '../util/exit-codes';

export type ErrorCode =
  | 'CONFIG_MISSING'
  | 'CONFIG_INVALID'
  | 'AUTH_REQUIRED'
  | 'ACCESS_DENIED'
  | 'NOT_FOUND'
  | 'LOCKED'
  | 'QUOTA_EXCEEDED'
  | 'UPSTREAM'
  | 'TIMEOUT'
  | 'IO';

const EXIT_BY_CODE: Record<ErrorCode, ExitCodeValue> = {
  CONFIG_MISSING: ExitCode.Config,
  CONFIG_INVALID: ExitCode.Config,
  AUTH_REQUIRED: ExitCode.AuthRequired,
  ACCESS_DENIED: ExitCode.Upstream,
  NOT_FOUND: ExitCode.Upstream,
  LOCKED: ExitCode.Upstream,
  QUOTA_EXCEEDED: ExitCode.Upstream,
  UPSTREAM: ExitCode.Upstream,
  TIMEOUT: ExitCode.Upstream,
  IO: ExitCode.Io,
};

export class CliError extends Error {
  constructor(
    public readonly code: ErrorCode,
    message: string,
    public readonly detail?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'CliError';
    Object.setPrototypeOf(this, new.target.prototype);
  }

  toExit(): ExitWithCode {
    return new ExitWithCode(EXIT_BY_CODE[this.code], {
      code: this.code.toLowerCase(),
      message: this.message,
      ...(this.detail ?? {}),
    });
  }
}
```

- [ ] **Step 5: Implement `src/config/load.ts`**

Host validation must accept `<tenant>.sharepoint.com` and `<tenant>-my.sharepoint.com` and reject anything else, so a typo cannot redirect requests with live cookies attached to an attacker-controlled host.

```typescript
// src/config/load.ts
import * as os from 'node:os';
import * as path from 'node:path';

import { CliError } from './errors';

export interface CliConfig {
  host: string;
  httpTimeoutMs: number;
  loginTimeoutMs: number;
  renewTimeoutMs: number;
  chromeChannel: string;
  sessionPath: string;
  profileDir: string;
  lockPath: string;
}

export interface ConfigOverrides {
  host?: string;
  httpTimeoutMs?: number;
  loginTimeoutMs?: number;
  renewTimeoutMs?: number;
  chromeChannel?: string;
  sessionPath?: string;
}

/** Defaults for the four plumbing settings. See the exception in CLAUDE.md. */
const DEFAULTS = {
  httpTimeoutMs: 30_000,
  loginTimeoutMs: 300_000,
  renewTimeoutMs: 30_000,
  chromeChannel: 'chrome',
} as const;

/**
 * Only `<label>.sharepoint.com` is accepted. Anchored at both ends so
 * "evil.com/x.sharepoint.com" and "x.sharepoint.com.evil.com" both fail:
 * the host is interpolated straight into request URLs that carry live
 * session cookies.
 */
const HOST_RE = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?\.sharepoint\.com$/i;

export function stateDir(): string {
  return path.join(os.homedir(), '.sharepoint-cli');
}

export function defaultSessionPath(): string {
  return path.join(stateDir(), 'session.json');
}

function resolveInt(
  flag: number | undefined,
  envRaw: string | undefined,
  envName: string,
  fallback: number,
): number {
  if (flag !== undefined) return flag;
  if (envRaw === undefined) return fallback;
  const n = Number.parseInt(envRaw, 10);
  if (!Number.isFinite(n) || n <= 0) {
    throw new CliError('CONFIG_INVALID', `${envName} must be a positive integer, got "${envRaw}"`);
  }
  return n;
}

export function loadConfig(
  overrides: ConfigOverrides = {},
  env: NodeJS.ProcessEnv = process.env,
): CliConfig {
  const host = overrides.host ?? env.SHAREPOINT_CLI_HOST;
  if (!host) {
    throw new CliError(
      'CONFIG_MISSING',
      'host is required: pass --host <tenant>.sharepoint.com or set SHAREPOINT_CLI_HOST. It has no default because it identifies which tenant is addressed.',
    );
  }
  if (!HOST_RE.test(host)) {
    throw new CliError(
      'CONFIG_INVALID',
      `host must be a single <label>.sharepoint.com hostname, got "${host}"`,
    );
  }

  const sessionPath =
    overrides.sessionPath ?? env.SHAREPOINT_CLI_SESSION_FILE ?? defaultSessionPath();

  return {
    host: host.toLowerCase(),
    httpTimeoutMs: resolveInt(
      overrides.httpTimeoutMs,
      env.SHAREPOINT_CLI_HTTP_TIMEOUT_MS,
      'SHAREPOINT_CLI_HTTP_TIMEOUT_MS',
      DEFAULTS.httpTimeoutMs,
    ),
    loginTimeoutMs: resolveInt(
      overrides.loginTimeoutMs,
      env.SHAREPOINT_CLI_LOGIN_TIMEOUT_MS,
      'SHAREPOINT_CLI_LOGIN_TIMEOUT_MS',
      DEFAULTS.loginTimeoutMs,
    ),
    renewTimeoutMs: resolveInt(
      overrides.renewTimeoutMs,
      env.SHAREPOINT_CLI_RENEW_TIMEOUT_MS,
      'SHAREPOINT_CLI_RENEW_TIMEOUT_MS',
      DEFAULTS.renewTimeoutMs,
    ),
    chromeChannel:
      overrides.chromeChannel ?? env.SHAREPOINT_CLI_CHROME_CHANNEL ?? DEFAULTS.chromeChannel,
    sessionPath,
    profileDir: path.join(stateDir(), 'playwright-profile'),
    lockPath: path.join(stateDir(), '.browser.lock'),
  };
}
```

- [ ] **Step 6: Implement `src/session/schema.ts` and `src/session/store.ts`**

`schema.ts` holds the type and pure parse/serialize, ported from `outlook-access/src/session/sharepoint-schema.ts` with the `defaultSharepointSessionPath` helper dropped (it lives in config now). `store.ts` holds the async IO. Both re-exported from `store.ts` so tests import one place.

```typescript
// src/session/schema.ts
export interface SharepointSession {
  version: 1;
  host: string;
  /** Optional: cookie-auth tenants emit none. Sent as Authorization when present. */
  bearer?: string;
  /** Serialized cookie header, e.g. "FedAuth=…; rtFa=…". */
  cookies: string;
  capturedAt: string;
  tokenExpiresAt: string;
}

export class SessionParseError extends Error {
  constructor(msg: string) {
    super(msg);
    this.name = 'SessionParseError';
  }
}

export function parseSession(json: string): SharepointSession {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch (e) {
    throw new SessionParseError(`Invalid JSON: ${(e as Error).message}`);
  }
  if (typeof raw !== 'object' || raw === null) {
    throw new SessionParseError('Expected JSON object');
  }
  const obj = raw as Record<string, unknown>;
  if (obj.version !== 1) {
    throw new SessionParseError(`Unsupported version: ${String(obj.version)}`);
  }
  for (const key of ['host', 'cookies', 'capturedAt', 'tokenExpiresAt']) {
    if (typeof obj[key] !== 'string' || (obj[key] as string).length === 0) {
      throw new SessionParseError(`Missing or invalid "${key}"`);
    }
  }
  if (obj.bearer !== undefined && typeof obj.bearer !== 'string') {
    throw new SessionParseError('Invalid "bearer" (must be a string when present)');
  }
  return obj as unknown as SharepointSession;
}

export function serializeSession(s: SharepointSession): string {
  return JSON.stringify(s, null, 2);
}
```

`store.ts` mirrors `saveSharepointSession` from `outlook-access`: mkdir 0700, write `<path>.tmp` at 0600, rename. Re-export everything from `schema.ts`.

- [ ] **Step 7: Run tests to verify they pass**

Run: `npm test`
Expected: PASS, config and session suites green.

- [ ] **Step 8: Register in `project-functions.MD` and commit**

Mark F-105 `built`. Add a `Config` row to Cross-cutting.

```bash
git add -A
git commit -m "feat(config): tenant-host validation and session store"
```

---

### Task 2: Path handling

The single most error-prone part of SPO REST, and the one most exposed to tenant data, since document names here are routinely Greek.

**Files:**

- Create: `src/sharepoint/paths.ts`
- Test: `test_scripts/paths.spec.ts`

**Interfaces:**

- Produces: `normalizeServerRelative(p: string): string`, `odataLiteral(p: string): string`, `folderApi(p: string): string`, `fileApi(p: string): string`, `splitParentLeaf(p: string): { parent: string; leaf: string }`, `PathError`

- [ ] **Step 1: Write the failing test**

```typescript
import { describe, expect, it } from 'vitest';
import {
  normalizeServerRelative,
  odataLiteral,
  folderApi,
  fileApi,
  splitParentLeaf,
  PathError,
} from '../src/sharepoint/paths';

describe('normalizeServerRelative', () => {
  it('adds a leading slash', () => {
    expect(normalizeServerRelative('Shared Documents')).toBe('/Shared Documents');
  });

  it('collapses duplicate slashes and strips a trailing one', () => {
    expect(normalizeServerRelative('//a///b/')).toBe('/a/b');
  });

  it('keeps the root as a bare slash', () => {
    expect(normalizeServerRelative('/')).toBe('/');
  });

  it('rejects parent traversal', () => {
    expect(() => normalizeServerRelative('/a/../../etc')).toThrowError(PathError);
  });

  it('rejects an absolute URL', () => {
    expect(() => normalizeServerRelative('https://evil.example.com/x')).toThrowError(PathError);
  });

  it('rejects a protocol-relative URL', () => {
    expect(() => normalizeServerRelative('//evil.example.com/x')).toThrowError(PathError);
  });

  it('rejects NUL and newline injection', () => {
    expect(() => normalizeServerRelative('/a�b')).toThrowError(PathError);
    expect(() => normalizeServerRelative('/a\nb')).toThrowError(PathError);
  });

  it('preserves Greek characters unchanged', () => {
    expect(normalizeServerRelative('/Έγγραφα/Φύλλα')).toBe('/Έγγραφα/Φύλλα');
  });
});

describe('odataLiteral', () => {
  it('doubles a single apostrophe', () => {
    expect(odataLiteral("/a/O'Brien.docx")).toBe("/a/O''Brien.docx");
  });

  it('doubles every apostrophe, not just the first', () => {
    expect(odataLiteral("/a/'x'.docx")).toBe("/a/''x''.docx");
  });

  it('leaves a path without apostrophes alone', () => {
    expect(odataLiteral('/a/b.docx')).toBe('/a/b.docx');
  });
});

describe('folderApi / fileApi', () => {
  it('percent-encodes the URI layer while the literal stays decoded', () => {
    const u = folderApi('/Shared Documents');
    expect(u).toContain('GetFolderByServerRelativePath');
    expect(u).toContain('%20');
    expect(u).not.toContain('Shared Documents');
  });

  it('percent-encodes Greek names', () => {
    const u = folderApi('/Έγγραφα');
    expect(u).toContain('%CE%88');
    expect(u).not.toContain('Έγγραφα');
  });

  it('escapes the apostrophe before percent-encoding, so both layers apply', () => {
    // %27%27 is the encoded form of the DOUBLED apostrophe. A single %27
    // would mean the OData escape was skipped and the literal terminates early.
    expect(fileApi("/a/O'B.docx")).toContain('%27%27');
  });

  it('uses the Path accessor, never the Url accessor', () => {
    expect(fileApi('/a/b')).toContain('GetFileByServerRelativePath');
    expect(fileApi('/a/b')).not.toContain('GetFileByServerRelativeUrl');
  });

  it('encodes a hash in a filename, which the Url accessor mishandles', () => {
    expect(fileApi('/a/b#1.docx')).toContain('%23');
  });
});

describe('splitParentLeaf', () => {
  it('splits a nested path', () => {
    expect(splitParentLeaf('/a/b/c.docx')).toEqual({ parent: '/a/b', leaf: 'c.docx' });
  });

  it('splits a top-level path to the root parent', () => {
    expect(splitParentLeaf('/c.docx')).toEqual({ parent: '/', leaf: 'c.docx' });
  });

  it('rejects the root, which has no leaf', () => {
    expect(() => splitParentLeaf('/')).toThrowError(PathError);
  });

  it('handles Greek leaves', () => {
    expect(splitParentLeaf('/Έγγραφα/α.docx')).toEqual({ parent: '/Έγγραφα', leaf: 'α.docx' });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test_scripts/paths.spec.ts`
Expected: FAIL, cannot resolve module.

- [ ] **Step 3: Implement `src/sharepoint/paths.ts`**

The two encoding layers are the whole point of this module. `odataLiteral` escapes for the OData string grammar (apostrophe doubling). `encodeURIComponent` then escapes for the URI. They must be applied in that order: encoding first would percent-encode the apostrophe so the doubling never matches, and the literal would terminate early on the decoded server side.

```typescript
// src/sharepoint/paths.ts
//
// Server-relative path handling for classic SPO REST.
//
// TWO ENCODING LAYERS, applied in this order and never conflated:
//   1. OData string literal  — a single quote must be doubled ('' )
//   2. URI percent-encoding  — applied to the already-escaped literal
// Reversing them percent-encodes the quote so the doubling never matches,
// and the literal terminates early once SharePoint decodes the URI.

export class PathError extends Error {
  constructor(msg: string) {
    super(msg);
    this.name = 'PathError';
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

/** Control characters, including NUL and newline, that must never reach a URL. */
// eslint-disable-next-line no-control-regex
const CONTROL_RE = /[�-]/;

export function normalizeServerRelative(input: string): string {
  if (typeof input !== 'string' || input.length === 0) {
    throw new PathError('path must be a non-empty string');
  }
  if (CONTROL_RE.test(input)) {
    throw new PathError('path contains control characters');
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(input)) {
    throw new PathError(`path must be server-relative, got an absolute URL: "${input}"`);
  }
  // "//host/x" is protocol-relative. Reject before collapsing slashes, since
  // collapsing would turn it into an innocuous-looking "/host/x".
  if (input.startsWith('//')) {
    throw new PathError(`path must be server-relative, got a protocol-relative URL: "${input}"`);
  }

  const withLead = input.startsWith('/') ? input : '/' + input;
  const collapsed = withLead.replace(/\/{2,}/g, '/');
  const segments = collapsed.split('/');
  for (const seg of segments) {
    if (seg === '..') throw new PathError(`path may not contain "..": "${input}"`);
  }
  const trimmed = collapsed.length > 1 ? collapsed.replace(/\/+$/, '') : collapsed;
  return trimmed.length === 0 ? '/' : trimmed;
}

/** Layer 1: escape for the OData string-literal grammar. */
export function odataLiteral(p: string): string {
  return p.split("'").join("''");
}

/** Layer 1 then layer 2, producing the value to interpolate into a REST URL. */
function encodedLiteral(p: string): string {
  return encodeURIComponent(odataLiteral(normalizeServerRelative(p)));
}

export function folderApi(p: string): string {
  return `/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodedLiteral(p)}')`;
}

export function fileApi(p: string): string {
  return `/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedLiteral(p)}')`;
}

export function splitParentLeaf(p: string): { parent: string; leaf: string } {
  const norm = normalizeServerRelative(p);
  if (norm === '/') throw new PathError('root path has no leaf name');
  const idx = norm.lastIndexOf('/');
  const parent = idx === 0 ? '/' : norm.slice(0, idx);
  const leaf = norm.slice(idx + 1);
  if (leaf.length === 0) throw new PathError(`path has no leaf name: "${p}"`);
  return { parent, leaf };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test_scripts/paths.spec.ts`
Expected: PASS, all cases green.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat(paths): server-relative path encoding with two-layer escaping"
```

---

### Task 3: HTTP errors and client

**Files:**

- Create: `src/http/errors.ts`, `src/http/client.ts`, `src/http/types.ts`
- Test: `test_scripts/http-errors.spec.ts`, `test_scripts/http-client.spec.ts`

**Interfaces:**

- Consumes: `CliError` from Task 1, `SharepointSession` from Task 1
- Produces: `SharepointHttpError { status, url, body, code }`, `classifyStatus(status, body): ErrorCode`, `isStaleDigest(status, body): boolean`, `SharepointClient` with `getJson<T>(path)`, `getBinary(pathOrUrl)`, `postJson<T>(path, body?, extraHeaders?)`, `postBinary<T>(path, bytes, extraHeaders?)`, and `setDigestProvider(fn)`

- [ ] **Step 1: Write the failing test**

```typescript
import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import { classifyStatus, isStaleDigest, SharepointHttpError } from '../src/http/errors';
import { SharepointClient } from '../src/http/client';

const STALE =
  '{"error":{"message":"The security validation for this page is invalid and might be corrupted."}}';

describe('isStaleDigest', () => {
  it('detects the digest marker on a 403', () => {
    expect(isStaleDigest(403, STALE)).toBe(true);
  });

  it('is case-insensitive', () => {
    expect(isStaleDigest(403, 'SECURITY VALIDATION FOR THIS PAGE IS INVALID')).toBe(true);
  });

  it('does not fire on a plain 403 access denial', () => {
    expect(isStaleDigest(403, '{"error":{"code":"accessDenied"}}')).toBe(false);
  });

  it('does not fire on a 401 even with the marker present', () => {
    expect(isStaleDigest(401, STALE)).toBe(false);
  });
});

describe('classifyStatus', () => {
  it.each([
    [401, '', 'AUTH_REQUIRED'],
    [403, '{"error":"accessDenied"}', 'ACCESS_DENIED'],
    [404, '', 'NOT_FOUND'],
    [423, '', 'LOCKED'],
    [507, '', 'QUOTA_EXCEEDED'],
    [500, '', 'UPSTREAM'],
    [418, '', 'UPSTREAM'],
  ])('maps %i to %s', (status, body, expected) => {
    expect(classifyStatus(status as number, body as string)).toBe(expected);
  });
});

describe('SharepointClient', () => {
  const session = {
    version: 1 as const,
    host: 'x.sharepoint.com',
    cookies: 'FedAuth=aaa',
    capturedAt: '2026-08-08T00:00:00.000Z',
    tokenExpiresAt: '2099-01-01T00:00:00.000Z',
  };
  let fetchMock: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    fetchMock = vi.fn();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  const ok = (body: unknown) =>
    new Response(JSON.stringify(body), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    });

  it('sends the Cookie header and no Authorization when there is no bearer', async () => {
    fetchMock.mockResolvedValue(ok({ Title: 'T' }));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await c.getJson('/_api/web');
    const headers = fetchMock.mock.calls[0][1].headers as Record<string, string>;
    expect(headers.Cookie).toBe('FedAuth=aaa');
    expect(headers.Authorization).toBeUndefined();
  });

  it('sends Authorization when a bearer is present', async () => {
    fetchMock.mockResolvedValue(ok({}));
    const c = new SharepointClient({ ...session, bearer: 'jwt' }, { httpTimeoutMs: 1000 });
    await c.getJson('/_api/web');
    const headers = fetchMock.mock.calls[0][1].headers as Record<string, string>;
    expect(headers.Authorization).toBe('Bearer jwt');
  });

  it('builds the URL from the session host', async () => {
    fetchMock.mockResolvedValue(ok({}));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await c.getJson('/_api/web');
    expect(fetchMock.mock.calls[0][0]).toBe('https://x.sharepoint.com/_api/web');
  });

  it('refetches the digest and retries once on a stale-digest 403', async () => {
    fetchMock
      .mockResolvedValueOnce(new Response(STALE, { status: 403 }))
      .mockResolvedValueOnce(ok({ ok: true }));
    const digest = vi.fn().mockResolvedValueOnce('D1').mockResolvedValueOnce('D2');
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    c.setDigestProvider(digest);

    const res = await c.postJson<{ ok: boolean }>('/_api/web/folders/addUsingPath');

    expect(res.ok).toBe(true);
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(digest).toHaveBeenNthCalledWith(1, false);
    expect(digest).toHaveBeenNthCalledWith(2, true); // forced refresh
    expect((fetchMock.mock.calls[1][1].headers as Record<string, string>)['X-RequestDigest']).toBe(
      'D2',
    );
  });

  it('does not retry twice on a repeated stale-digest 403', async () => {
    fetchMock.mockResolvedValue(new Response(STALE, { status: 403 }));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    c.setDigestProvider(vi.fn().mockResolvedValue('D'));
    await expect(c.postJson('/_api/web/x')).rejects.toThrowError(SharepointHttpError);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('does not retry a plain 403 access denial', async () => {
    fetchMock.mockResolvedValue(new Response('{"error":"accessDenied"}', { status: 403 }));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    c.setDigestProvider(vi.fn().mockResolvedValue('D'));
    await expect(c.postJson('/_api/web/x')).rejects.toMatchObject({ code: 'ACCESS_DENIED' });
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('never retries a GET, which carries no digest', async () => {
    fetchMock.mockResolvedValue(new Response(STALE, { status: 403 }));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await expect(c.getJson('/_api/web')).rejects.toThrowError(SharepointHttpError);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('rejects a getBinary absolute URL on a foreign host', async () => {
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await expect(c.getBinary('https://evil.example.com/a.docx')).rejects.toThrowError(/host/i);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('allows a getBinary absolute URL on a sharepoint.com host', async () => {
    fetchMock.mockResolvedValue(new Response('bytes', { status: 200 }));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await c.getBinary('https://x-my.sharepoint.com/personal/a/Doc.docx');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('does not put the bearer or cookies into the error message', async () => {
    fetchMock.mockResolvedValue(new Response('nope', { status: 500 }));
    const c = new SharepointClient({ ...session, bearer: 'SECRETJWT' }, { httpTimeoutMs: 1000 });
    const err = await c.getJson('/_api/web').catch((e) => e as Error);
    expect(err.message).not.toContain('SECRETJWT');
    expect(err.message).not.toContain('FedAuth');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test_scripts/http-errors.spec.ts test_scripts/http-client.spec.ts`
Expected: FAIL, cannot resolve modules.

- [ ] **Step 3: Implement `src/http/errors.ts`**

```typescript
// src/http/errors.ts
import type { ErrorCode } from '../config/errors';

/**
 * SharePoint answers a STALE DIGEST with 403, not 401. Without this
 * discriminator every expired write looks like a permissions failure and
 * sends the operator down the wrong diagnostic path.
 */
const STALE_DIGEST_MARKER = 'security validation for this page is invalid';

export function isStaleDigest(status: number, body: string): boolean {
  return status === 403 && body.toLowerCase().includes(STALE_DIGEST_MARKER);
}

export function classifyStatus(status: number, body: string): ErrorCode {
  if (status === 401) return 'AUTH_REQUIRED';
  if (status === 403) return isStaleDigest(status, body) ? 'UPSTREAM' : 'ACCESS_DENIED';
  if (status === 404 || status === 410) return 'NOT_FOUND';
  if (status === 423) return 'LOCKED';
  if (status === 507) return 'QUOTA_EXCEEDED';
  return 'UPSTREAM';
}

export class SharepointHttpError extends Error {
  public readonly code: ErrorCode;

  constructor(
    public readonly status: number,
    public readonly url: string,
    public readonly body: string,
  ) {
    // Message carries status and URL only. Never the session: callers log this.
    super(`SharePoint ${status} for ${url}`);
    this.name = 'SharepointHttpError';
    this.code = classifyStatus(status, body);
    Object.setPrototypeOf(this, new.target.prototype);
  }
}
```

- [ ] **Step 4: Implement `src/http/client.ts`**

Two security-relevant rules live here. Absolute URLs passed to `getBinary` must be host-checked before the cookie header is attached, otherwise a crafted attachment URL exfiltrates the session. And error messages must never interpolate the session.

```typescript
// src/http/client.ts
import type { SharepointSession } from '../session/schema';
import { CliError } from '../config/errors';
import { SharepointHttpError, isStaleDigest } from './errors';

export interface ClientOpts {
  httpTimeoutMs: number;
}

export type DigestProvider = (force: boolean) => Promise<string>;

/** Only *.sharepoint.com absolute URLs may receive the session cookies. */
function assertSharepointUrl(raw: string): URL {
  let u: URL;
  try {
    u = new URL(raw);
  } catch {
    throw new CliError('CONFIG_INVALID', `not a valid URL: "${raw}"`);
  }
  if (u.protocol !== 'https:') {
    throw new CliError('CONFIG_INVALID', `refusing non-https URL: "${raw}"`);
  }
  const h = u.hostname.toLowerCase();
  if (h !== 'sharepoint.com' && !h.endsWith('.sharepoint.com')) {
    throw new CliError(
      'CONFIG_INVALID',
      `refusing to send session cookies to non-SharePoint host "${u.hostname}"`,
    );
  }
  return u;
}

export class SharepointClient {
  private digestProvider: DigestProvider | null = null;

  constructor(
    private readonly session: SharepointSession,
    private readonly opts: ClientOpts,
  ) {}

  setDigestProvider(fn: DigestProvider): void {
    this.digestProvider = fn;
  }

  private baseHeaders(accept: string): Record<string, string> {
    const h: Record<string, string> = { Accept: accept };
    if (this.session.bearer) h.Authorization = `Bearer ${this.session.bearer}`;
    if (this.session.cookies) h.Cookie = this.session.cookies;
    return h;
  }

  private url(pathOrUrl: string): string {
    if (/^https?:/i.test(pathOrUrl)) return assertSharepointUrl(pathOrUrl).toString();
    return `https://${this.session.host}${pathOrUrl}`;
  }

  private async raw(
    method: 'GET' | 'POST',
    pathOrUrl: string,
    accept: string,
    body?: BodyInit,
    extra?: Record<string, string>,
  ): Promise<Response> {
    const url = this.url(pathOrUrl);
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), this.opts.httpTimeoutMs);
    try {
      return await fetch(url, {
        method,
        headers: { ...this.baseHeaders(accept), ...(extra ?? {}) },
        body,
        signal: ctrl.signal,
        redirect: 'follow',
      });
    } catch (err) {
      if ((err as Error).name === 'AbortError') {
        throw new CliError(
          'TIMEOUT',
          `request timed out after ${this.opts.httpTimeoutMs}ms: ${url}`,
        );
      }
      throw new CliError('UPSTREAM', `request failed: ${(err as Error).message}`);
    } finally {
      clearTimeout(timer);
    }
  }

  async getJson<T>(path: string): Promise<T> {
    const resp = await this.raw('GET', path, 'application/json;odata=nometadata');
    if (!resp.ok) throw new SharepointHttpError(resp.status, this.url(path), await resp.text());
    return (await resp.json()) as T;
  }

  async getBinary(
    pathOrUrl: string,
  ): Promise<{ bytes: Buffer; contentType: string; filename?: string }> {
    const resp = await this.raw('GET', pathOrUrl, '*/*');
    if (!resp.ok) {
      throw new SharepointHttpError(resp.status, this.url(pathOrUrl), await resp.text());
    }
    return {
      bytes: Buffer.from(await resp.arrayBuffer()),
      contentType: resp.headers.get('content-type') ?? 'application/octet-stream',
      filename: parseContentDispositionFilename(resp.headers.get('content-disposition')),
    };
  }

  /**
   * POST with the retry-once-on-stale-digest envelope. Mirrors the 401
   * envelope in outlook-access/src/http/outlook-client.ts:435 so all three
   * CLIs behave the same under credential churn.
   */
  private async post(
    path: string,
    body: BodyInit | undefined,
    extra: Record<string, string>,
    accept: string,
  ): Promise<Response> {
    const send = async (force: boolean): Promise<Response> => {
      const headers = { ...extra };
      if (this.digestProvider) headers['X-RequestDigest'] = await this.digestProvider(force);
      return this.raw('POST', path, accept, body, headers);
    };

    const first = await send(false);
    if (first.ok) return first;

    const text = await first.text();
    if (isStaleDigest(first.status, text) && this.digestProvider) {
      const second = await send(true);
      if (second.ok) return second;
      throw new SharepointHttpError(second.status, this.url(path), await second.text());
    }
    throw new SharepointHttpError(first.status, this.url(path), text);
  }

  async postJson<T>(path: string, body?: unknown, extra: Record<string, string> = {}): Promise<T> {
    const payload = body === undefined ? undefined : JSON.stringify(body);
    const headers =
      payload === undefined
        ? extra
        : { ...extra, 'Content-Type': 'application/json;odata=verbose' };
    const resp = await this.post(path, payload, headers, 'application/json;odata=nometadata');
    const text = await resp.text();
    return (text.length > 0 ? JSON.parse(text) : {}) as T;
  }

  async postBinary<T>(path: string, bytes: Buffer, extra: Record<string, string> = {}): Promise<T> {
    const resp = await this.post(
      path,
      new Uint8Array(bytes),
      extra,
      'application/json;odata=nometadata',
    );
    const text = await resp.text();
    return (text.length > 0 ? JSON.parse(text) : {}) as T;
  }
}

function parseContentDispositionFilename(header: string | null): string | undefined {
  if (!header) return undefined;
  const m1 = header.match(/filename\*=(?:UTF-8'')?([^;]+)/i);
  if (m1) {
    try {
      return decodeURIComponent(m1[1].trim().replace(/^"|"$/g, ''));
    } catch {
      /* fall through */
    }
  }
  const m2 = header.match(/filename=("?)([^";]+)\1/i);
  return m2 ? m2[2].trim() : undefined;
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `npx vitest run test_scripts/http-errors.spec.ts test_scripts/http-client.spec.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat(http): REST client with stale-digest discrimination and host pinning"
```

---

### Task 4: Request digest

**Files:**

- Create: `src/sharepoint/digest.ts`
- Test: `test_scripts/digest.spec.ts`

**Interfaces:**

- Consumes: `SharepointClient.postJson` from Task 3
- Produces: `class DigestCache { constructor(client: SharepointClient, now?: () => number); get(force?: boolean): Promise<string> }`

- [ ] **Step 1: Write the failing test**

```typescript
import { describe, expect, it, vi } from 'vitest';
import { DigestCache } from '../src/sharepoint/digest';

function clientReturning(
  ...values: Array<{ FormDigestValue: string; FormDigestTimeoutSeconds: number }>
) {
  const postJson = vi.fn();
  for (const v of values) postJson.mockResolvedValueOnce(v);
  return { postJson } as never;
}

describe('DigestCache', () => {
  it('fetches once and caches', async () => {
    const c = clientReturning({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 });
    const d = new DigestCache(c);
    expect(await d.get()).toBe('D1');
    expect(await d.get()).toBe('D1');
    expect((c as unknown as { postJson: ReturnType<typeof vi.fn> }).postJson).toHaveBeenCalledTimes(
      1,
    );
  });

  it('calls the contextinfo endpoint', async () => {
    const c = clientReturning({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 });
    await new DigestCache(c).get();
    expect((c as unknown as { postJson: ReturnType<typeof vi.fn> }).postJson).toHaveBeenCalledWith(
      '/_api/contextinfo',
    );
  });

  it('refetches after expiry, applying the 60s safety margin', async () => {
    const c = clientReturning(
      { FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 },
      { FormDigestValue: 'D2', FormDigestTimeoutSeconds: 1800 },
    );
    let t = 1_000_000;
    const d = new DigestCache(c, () => t);
    expect(await d.get()).toBe('D1');
    // 1740s = 1800 - 60. One ms before that boundary the cache must still hold.
    t += 1_739_000;
    expect(await d.get()).toBe('D1');
    t += 2_000;
    expect(await d.get()).toBe('D2');
  });

  it('refetches when forced, even inside the validity window', async () => {
    const c = clientReturning(
      { FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 },
      { FormDigestValue: 'D2', FormDigestTimeoutSeconds: 1800 },
    );
    const d = new DigestCache(c, () => 1_000_000);
    expect(await d.get()).toBe('D1');
    expect(await d.get(true)).toBe('D2');
  });

  it('treats a missing timeout as a conservative 60s of validity', async () => {
    const c = clientReturning({ FormDigestValue: 'D1' } as never, {
      FormDigestValue: 'D2',
      FormDigestTimeoutSeconds: 1800,
    });
    let t = 1_000_000;
    const d = new DigestCache(c, () => t);
    expect(await d.get()).toBe('D1');
    t += 61_000;
    expect(await d.get()).toBe('D2');
  });

  it('does not issue concurrent fetches for simultaneous callers', async () => {
    const postJson = vi
      .fn()
      .mockImplementation(
        () =>
          new Promise((r) =>
            setTimeout(() => r({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 }), 10),
          ),
      );
    const d = new DigestCache({ postJson } as never);
    const [a, b] = await Promise.all([d.get(), d.get()]);
    expect(a).toBe('D1');
    expect(b).toBe('D1');
    expect(postJson).toHaveBeenCalledTimes(1);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test_scripts/digest.spec.ts`
Expected: FAIL, cannot resolve module.

- [ ] **Step 3: Implement `src/sharepoint/digest.ts`**

```typescript
// src/sharepoint/digest.ts
//
// X-RequestDigest cache. Process-local and never persisted: it is a CSRF
// token whose lifetime (1800s) is shorter than most cron gaps, so writing it
// to disk would add attack surface for nothing.

import type { SharepointClient } from '../http/client';

interface ContextInfo {
  FormDigestValue?: string;
  FormDigestTimeoutSeconds?: number;
}

/** Applied when SharePoint omits the timeout. Deliberately pessimistic. */
const FALLBACK_TTL_S = 60;
/** Covers clock skew and uploads that start just before expiry. */
const SAFETY_MARGIN_MS = 60_000;

export class DigestCache {
  private value: string | null = null;
  private expiresAtMs = 0;
  private inflight: Promise<string> | null = null;

  constructor(
    private readonly client: Pick<SharepointClient, 'postJson'>,
    private readonly now: () => number = () => Date.now(),
  ) {}

  async get(force = false): Promise<string> {
    if (!force && this.value !== null && this.now() < this.expiresAtMs) {
      return this.value;
    }
    // Collapse concurrent callers onto one request.
    if (!force && this.inflight) return this.inflight;

    const p = this.fetch();
    this.inflight = p;
    try {
      return await p;
    } finally {
      if (this.inflight === p) this.inflight = null;
    }
  }

  private async fetch(): Promise<string> {
    const info = await this.client.postJson<ContextInfo>('/_api/contextinfo');
    const digest = info?.FormDigestValue;
    if (typeof digest !== 'string' || digest.length === 0) {
      throw new Error('contextinfo returned no FormDigestValue');
    }
    const ttlS =
      typeof info.FormDigestTimeoutSeconds === 'number' && info.FormDigestTimeoutSeconds > 0
        ? info.FormDigestTimeoutSeconds
        : FALLBACK_TTL_S;
    this.value = digest;
    this.expiresAtMs = this.now() + ttlS * 1000 - SAFETY_MARGIN_MS;
    return digest;
  }
}
```

Note the fallback interaction: with `FALLBACK_TTL_S = 60` and a 60s margin the expiry lands exactly at `now`, so the next call refetches. That is the intended pessimism when SharePoint omits the field.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test_scripts/digest.spec.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat(digest): X-RequestDigest cache with safety margin and single-flight"
```

---

### Task 5: Read commands

**Files:**

- Create: `src/http/types.ts`, `src/commands/ls.ts`, `src/commands/get.ts`, `src/commands/libraries.ts`, `src/commands/search.ts`
- Test: `test_scripts/commands-read.spec.ts`

**Interfaces:**

- Consumes: `SharepointClient`, `folderApi`, `fileApi`, `normalizeServerRelative`
- Produces: `runLs(client, path): Promise<LsResult>` where `LsResult = { path: string; folders: Entry[]; files: Entry[] }` and `Entry = { name: string; serverRelativeUrl: string; size?: number; modified?: string }`; `runGet(client, pathOrUrl, outPath?): Promise<GetResult>`; `runLibraries(client): Promise<LibrariesResult>`; `runSearch(client, query, rows): Promise<SearchResult>`

- [ ] **Step 1: Write the failing test**

```typescript
import { describe, expect, it, vi } from 'vitest';
import { runLs } from '../src/commands/ls';
import { runLibraries } from '../src/commands/libraries';
import { runSearch } from '../src/commands/search';
import { runGet } from '../src/commands/get';

const client = (getJson: unknown, getBinary?: unknown) => ({ getJson, getBinary }) as never;

describe('runLs', () => {
  it('expands Folders and Files in one request', async () => {
    const getJson = vi.fn().mockResolvedValue({ Folders: [], Files: [] });
    await runLs(client(getJson), '/Έγγραφα');
    expect(getJson).toHaveBeenCalledTimes(1);
    const url = getJson.mock.calls[0][0] as string;
    expect(url).toContain('GetFolderByServerRelativePath');
    expect(url).toContain('$expand=Folders,Files');
  });

  it('maps folders and files into a stable shape', async () => {
    const getJson = vi.fn().mockResolvedValue({
      Folders: [
        { Name: 'Sub', ServerRelativeUrl: '/a/Sub', TimeLastModified: '2026-01-01T00:00:00Z' },
      ],
      Files: [
        {
          Name: 'α.docx',
          ServerRelativeUrl: '/a/α.docx',
          Length: '42',
          TimeLastModified: '2026-01-02T00:00:00Z',
        },
      ],
    });
    const r = await runLs(client(getJson), '/a');
    expect(r.folders).toEqual([
      { name: 'Sub', serverRelativeUrl: '/a/Sub', modified: '2026-01-01T00:00:00Z' },
    ]);
    expect(r.files).toEqual([
      {
        name: 'α.docx',
        serverRelativeUrl: '/a/α.docx',
        size: 42,
        modified: '2026-01-02T00:00:00Z',
      },
    ]);
  });

  it('coerces the string Length SharePoint returns into a number', async () => {
    const getJson = vi.fn().mockResolvedValue({
      Folders: [],
      Files: [{ Name: 'f', ServerRelativeUrl: '/f', Length: '1048576' }],
    });
    const r = await runLs(client(getJson), '/');
    expect(r.files[0].size).toBe(1048576);
  });

  it('tolerates a response with neither collection present', async () => {
    const getJson = vi.fn().mockResolvedValue({});
    const r = await runLs(client(getJson), '/a');
    expect(r.folders).toEqual([]);
    expect(r.files).toEqual([]);
  });
});

describe('runLibraries', () => {
  it('filters to visible document libraries', async () => {
    const getJson = vi.fn().mockResolvedValue({ value: [] });
    await runLibraries(client(getJson));
    const url = getJson.mock.calls[0][0] as string;
    expect(url).toContain('BaseTemplate%20eq%20101');
    expect(url).toContain('Hidden%20eq%20false');
  });

  it('maps the root folder URL out of the expansion', async () => {
    const getJson = vi.fn().mockResolvedValue({
      value: [{ Title: 'Έγγραφα', Id: 'guid-1', RootFolder: { ServerRelativeUrl: '/Έγγραφα' } }],
    });
    const r = await runLibraries(client(getJson));
    expect(r.libraries).toEqual([
      { title: 'Έγγραφα', id: 'guid-1', serverRelativeUrl: '/Έγγραφα' },
    ]);
  });
});

describe('runSearch', () => {
  it('quotes the query and passes rowlimit', async () => {
    const getJson = vi
      .fn()
      .mockResolvedValue({ PrimaryQueryResult: { RelevantResults: { Table: { Rows: [] } } } });
    await runSearch(client(getJson), 'budget', 5);
    const url = getJson.mock.calls[0][0] as string;
    expect(url).toContain("querytext='budget'");
    expect(url).toContain('rowlimit=5');
  });

  it('escapes an apostrophe in the query rather than breaking the literal', async () => {
    const getJson = vi
      .fn()
      .mockResolvedValue({ PrimaryQueryResult: { RelevantResults: { Table: { Rows: [] } } } });
    await runSearch(client(getJson), "O'Brien", 5);
    expect(getJson.mock.calls[0][0] as string).toContain("O''Brien");
  });

  it('flattens the Cells key/value rows into objects', async () => {
    const getJson = vi.fn().mockResolvedValue({
      PrimaryQueryResult: {
        RelevantResults: {
          Table: {
            Rows: [
              {
                Cells: [
                  { Key: 'Title', Value: 'T' },
                  { Key: 'Path', Value: 'https://x/a' },
                ],
              },
            ],
          },
        },
      },
    });
    const r = await runSearch(client(getJson), 'q', 1);
    expect(r.results).toEqual([{ Title: 'T', Path: 'https://x/a' }]);
  });

  it('rejects a non-positive rowlimit', async () => {
    await expect(runSearch(client(vi.fn()), 'q', 0)).rejects.toThrowError(/rows/);
  });
});

describe('runGet', () => {
  it('appends /$value to the file accessor', async () => {
    const getBinary = vi
      .fn()
      .mockResolvedValue({ bytes: Buffer.from('x'), contentType: 'text/plain' });
    await runGet(client(vi.fn(), getBinary), '/a/b.txt');
    expect(getBinary.mock.calls[0][0] as string).toMatch(/GetFileByServerRelativePath.*\/\$value$/);
  });

  it('passes an absolute URL straight through for host checking downstream', async () => {
    const getBinary = vi
      .fn()
      .mockResolvedValue({ bytes: Buffer.from('x'), contentType: 'text/plain' });
    await runGet(client(vi.fn(), getBinary), 'https://x.sharepoint.com/a.docx');
    expect(getBinary.mock.calls[0][0]).toBe('https://x.sharepoint.com/a.docx');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test_scripts/commands-read.spec.ts`
Expected: FAIL, cannot resolve modules.

- [ ] **Step 3: Implement the four command modules**

Each is a pure function taking a client and returning a plain object, so the CLI layer only formats. `runSearch` must apply `odataLiteral` to the user query for the same reason paths do: an apostrophe otherwise terminates the OData literal.

Key URLs:

```typescript
// ls.ts
`${folderApi(p)}?$expand=Folders,Files&$select=Name,ServerRelativeUrl,TimeLastModified,Folders/Name,Folders/ServerRelativeUrl,Folders/TimeLastModified,Files/Name,Files/ServerRelativeUrl,Files/Length,Files/TimeLastModified`
// get.ts
`${fileApi(p)}/$value` // or the absolute URL unchanged
// libraries.ts
`/_api/web/lists?$filter=${encodeURIComponent('BaseTemplate eq 101 and Hidden eq false')}&$select=Title,Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`
// search.ts
`/_api/search/query?querytext='${encodeURIComponent(odataLiteral(query))}'&rowlimit=${rows}&selectproperties='${encodeURIComponent('Title,Path,FileType,LastModifiedTime,Size')}'`;
```

`Length` arrives as a string and must be coerced with `Number.parseInt`. `runSearch` validates `rows` is a positive integer before building the URL.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test_scripts/commands-read.spec.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "feat(read): ls, get, libraries and search commands"
```

---

### Task 6: Upload and write commands

**Files:**

- Create: `src/sharepoint/upload.ts`, `src/commands/mkdir.ts`, `src/commands/put.ts`
- Test: `test_scripts/upload.spec.ts`, `test_scripts/commands-write.spec.ts`

**Interfaces:**

- Consumes: `SharepointClient.postBinary/postJson`, `splitParentLeaf`, `folderApi`, `fileApi`
- Produces: `CHUNK_THRESHOLD_BYTES`, `CHUNK_SIZE_BYTES`, `uploadFile(client, bytes, remoteFolder, name, overwrite): Promise<UploadResult>`; `runMkdir(client, path)`; `runPut(client, localPath, remoteFolder, overwrite)`

- [ ] **Step 1: Write the failing test**

```typescript
import { describe, expect, it, vi } from 'vitest';
import { uploadFile, CHUNK_THRESHOLD_BYTES, CHUNK_SIZE_BYTES } from '../src/sharepoint/upload';

function mkClient() {
  return {
    postBinary: vi.fn().mockResolvedValue({ ServerRelativeUrl: '/a/f.bin' }),
    postJson: vi.fn().mockResolvedValue({}),
  };
}

describe('uploadFile', () => {
  it('uses a single addUsingPath request below the threshold', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(1024), '/a', 'f.bin', true);
    expect(c.postBinary).toHaveBeenCalledTimes(1);
    const url = c.postBinary.mock.calls[0][0] as string;
    expect(url).toContain('addUsingPath');
    expect(url).toContain('overwrite=true');
    expect(url).not.toContain('StartUpload');
  });

  it('never uses the legacy add(url=) form', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(16), '/a', 'f.bin', false);
    expect(c.postBinary.mock.calls[0][0] as string).not.toMatch(/\/add\(url=/);
  });

  it('percent-encodes a Greek leaf name', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(16), '/a', 'α.docx', true);
    expect(c.postBinary.mock.calls[0][0] as string).toContain('%CE%B1');
  });

  it('chunks at exactly the threshold', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_THRESHOLD_BYTES), '/a', 'big.bin', true);
    const urls = c.postBinary.mock.calls.map((x) => x[0] as string);
    expect(urls.some((u) => u.includes('StartUpload'))).toBe(true);
    expect(urls.some((u) => u.includes('FinishUpload'))).toBe(true);
  });

  it('creates the empty file first, then starts the upload session', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_THRESHOLD_BYTES), '/a', 'big.bin', true);
    const urls = c.postBinary.mock.calls.map((x) => x[0] as string);
    expect(urls[0]).toContain('addUsingPath');
    expect(urls[1]).toContain('StartUpload');
  });

  it('splits an exact multiple into Start plus Continues plus Finish with no empty final chunk', async () => {
    const c = mkClient();
    const size = CHUNK_SIZE_BYTES * 3;
    await uploadFile(c as never, Buffer.alloc(size), '/a', 'big.bin', true);
    const urls = c.postBinary.mock.calls.map((x) => x[0] as string);
    expect(urls.filter((u) => u.includes('StartUpload')).length).toBe(1);
    expect(urls.filter((u) => u.includes('ContinueUpload')).length).toBe(1);
    expect(urls.filter((u) => u.includes('FinishUpload')).length).toBe(1);
  });

  it('sends correct fileOffset values', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 2 + 5), '/a', 'big.bin', true);
    const urls = c.postBinary.mock.calls.map((x) => x[0] as string);
    expect(urls.find((u) => u.includes('ContinueUpload'))).toContain(
      `fileOffset=${CHUNK_SIZE_BYTES}`,
    );
    expect(urls.find((u) => u.includes('FinishUpload'))).toContain(
      `fileOffset=${CHUNK_SIZE_BYTES * 2}`,
    );
  });

  it('uses one uploadId GUID for the whole session', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 2 + 1), '/a', 'big.bin', true);
    const ids = c.postBinary.mock.calls
      .map((x) => (x[0] as string).match(/uploadId=guid'([^']+)'/)?.[1])
      .filter(Boolean);
    expect(ids.length).toBeGreaterThan(1);
    expect(new Set(ids).size).toBe(1);
  });

  it('deletes the partial file when a chunk fails, then rethrows', async () => {
    const c = mkClient();
    c.postBinary
      .mockResolvedValueOnce({}) // addUsingPath
      .mockResolvedValueOnce({}) // StartUpload
      .mockRejectedValueOnce(new Error('network died')); // ContinueUpload
    await expect(
      uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 3), '/a', 'big.bin', true),
    ).rejects.toThrowError('network died');
    expect(c.postJson).toHaveBeenCalledWith(
      expect.stringContaining('recycle'),
      undefined,
      expect.anything(),
    );
  });

  it('does not mask the original error if cleanup also fails', async () => {
    const c = mkClient();
    c.postBinary
      .mockResolvedValueOnce({})
      .mockResolvedValueOnce({})
      .mockRejectedValueOnce(new Error('network died'));
    c.postJson.mockRejectedValue(new Error('cleanup failed too'));
    await expect(
      uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 3), '/a', 'big.bin', true),
    ).rejects.toThrowError('network died');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test_scripts/upload.spec.ts`
Expected: FAIL, cannot resolve module.

- [ ] **Step 3: Implement `src/sharepoint/upload.ts`**

```typescript
// src/sharepoint/upload.ts
import { randomUUID } from 'node:crypto';

import type { SharepointClient } from '../http/client';
import { folderApi, fileApi, normalizeServerRelative, odataLiteral } from './paths';

/** At or above this size, use a chunked session. */
export const CHUNK_THRESHOLD_BYTES = 10 * 1024 * 1024;
export const CHUNK_SIZE_BYTES = 10 * 1024 * 1024;

export interface UploadResult {
  serverRelativeUrl: string;
  size: number;
  chunked: boolean;
}

type Client = Pick<SharepointClient, 'postBinary' | 'postJson'>;

function addUsingPath(folder: string, leaf: string, overwrite: boolean): string {
  const encLeaf = encodeURIComponent(odataLiteral(leaf));
  return `${folderApi(folder)}/Files/addUsingPath(DecodedUrl='${encLeaf}',overwrite=${overwrite})`;
}

export async function uploadFile(
  client: Client,
  bytes: Buffer,
  remoteFolder: string,
  leaf: string,
  overwrite: boolean,
): Promise<UploadResult> {
  const folder = normalizeServerRelative(remoteFolder);
  const target = normalizeServerRelative(`${folder === '/' ? '' : folder}/${leaf}`);

  if (bytes.length < CHUNK_THRESHOLD_BYTES) {
    await client.postBinary(addUsingPath(folder, leaf, overwrite), bytes);
    return { serverRelativeUrl: target, size: bytes.length, chunked: false };
  }

  // SPO requires the target to exist before a chunked session starts.
  await client.postBinary(addUsingPath(folder, leaf, overwrite), Buffer.alloc(0));

  const uploadId = randomUUID();
  const file = fileApi(target);
  try {
    let offset = 0;
    // First chunk opens the session.
    const first = bytes.subarray(0, CHUNK_SIZE_BYTES);
    await client.postBinary(`${file}/StartUpload(uploadId=guid'${uploadId}')`, Buffer.from(first));
    offset = first.length;

    // Middle chunks. The final chunk always goes to FinishUpload, so stop
    // short of it here even when the size divides evenly.
    while (bytes.length - offset > CHUNK_SIZE_BYTES) {
      const chunk = bytes.subarray(offset, offset + CHUNK_SIZE_BYTES);
      await client.postBinary(
        `${file}/ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`,
        Buffer.from(chunk),
      );
      offset += chunk.length;
    }

    await client.postBinary(
      `${file}/FinishUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`,
      Buffer.from(bytes.subarray(offset)),
    );
    return { serverRelativeUrl: target, size: bytes.length, chunked: true };
  } catch (err) {
    // Leave no partial file behind: a retry must start clean rather than
    // resume into an inconsistent file.
    try {
      await client.postJson(`${file}/recycle()`, undefined, {});
    } catch {
      // Cleanup is best-effort and must never mask the original failure.
    }
    throw err;
  }
}
```

- [ ] **Step 4: Implement `src/commands/mkdir.ts` and `src/commands/put.ts`**

`runMkdir` posts to `/_api/web/folders/addUsingPath(DecodedUrl='<encoded>')`. `runPut` reads the local file, derives the leaf from its basename, and delegates to `uploadFile`. Neither invents parent folders.

`commands-write.spec.ts` covers: mkdir uses `addUsingPath` and rejects the root; put derives the leaf from the local basename; put propagates `overwrite`; put rejects a missing local file with `IO`.

- [ ] **Step 5: Run tests to verify they pass**

Run: `npx vitest run test_scripts/upload.spec.ts test_scripts/commands-write.spec.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat(write): chunked upload, mkdir and put"
```

---

### Task 7: Auth capture, login and renew

Ported from `outlook-access/src/auth/sharepoint-capture.ts`, made standalone. The substantive change is cold-profile support: the existing code assumes a warm context whose Microsoft SSO cookies were just set by an Outlook sign-in.

**Files:**

- Create: `src/auth/lock.ts`, `src/auth/capture.ts`, `src/commands/login.ts`, `src/commands/auth-renew.ts`
- Test: `test_scripts/lock.spec.ts`, `test_scripts/capture.spec.ts`

**Interfaces:**

- Produces: `acquireLock(path): Promise<() => Promise<void>>`; `isSharepointBearerUrl(host, url): boolean`; `deriveTokenExpiry(bearer, cookies): string`; `collectCookieHeader(all, host): string`; `captureSession(opts): Promise<SharepointSession>` where `opts = { host, profileDir, chromeChannel, timeoutMs, headless }`

- [ ] **Step 1: Port `src/auth/lock.ts` verbatim**

Copy `outlook-access/src/auth/lock.ts` unchanged except the error text, which becomes `another sharepoint-cli instance holds the lock: <path>`. It is already correct: `O_CREAT|O_EXCL` at 0600, stale-PID detection via signal 0, `EPERM` treated as alive, idempotent release.

- [ ] **Step 2: Write the failing tests**

`lock.spec.ts`: acquires in an empty dir and creates the file 0600; a second acquire while held throws; release is idempotent; a lock file holding a dead PID is treated as stale and reclaimed; a lock file with garbage content is reclaimed; release removes the file.

`capture.spec.ts` covers the pure helpers, which is where the real logic is:

```typescript
import { describe, expect, it } from 'vitest';
import { isSharepointBearerUrl, deriveTokenExpiry, collectCookieHeader } from '../src/auth/capture';

describe('isSharepointBearerUrl', () => {
  const host = 'x.sharepoint.com';
  it('matches a direct request to the host', () => {
    expect(isSharepointBearerUrl(host, 'https://x.sharepoint.com/_api/web')).toBe(true);
  });
  it('matches the MCAS-proxied rewrite', () => {
    expect(isSharepointBearerUrl(host, 'https://x-sharepoint-com.eu2.mcas.ms/_api/web')).toBe(true);
  });
  it('ignores an unrelated host', () => {
    expect(isSharepointBearerUrl(host, 'https://login.microsoftonline.com/x')).toBe(false);
  });
  it('ignores a lookalike host with the tenant as a path segment', () => {
    expect(isSharepointBearerUrl(host, 'https://evil.example.com/x.sharepoint.com/_api')).toBe(
      false,
    );
  });
});

describe('deriveTokenExpiry', () => {
  it('prefers the JWT exp when a bearer is present', () => {
    const exp = Math.floor(Date.UTC(2030, 0, 1) / 1000);
    const jwt = ['e30', Buffer.from(JSON.stringify({ exp })).toString('base64url'), 'sig'].join(
      '.',
    );
    expect(deriveTokenExpiry(jwt, [])).toBe(new Date(exp * 1000).toISOString());
  });
  it('falls back to FedAuth expiry when there is no bearer', () => {
    const secs = Math.floor(Date.UTC(2029, 0, 1) / 1000);
    expect(deriveTokenExpiry(undefined, [{ name: 'FedAuth', expires: secs }])).toBe(
      new Date(secs * 1000).toISOString(),
    );
  });
  it('prefers FedAuth over rtFa', () => {
    const a = Math.floor(Date.UTC(2029, 0, 1) / 1000);
    const b = Math.floor(Date.UTC(2028, 0, 1) / 1000);
    expect(
      deriveTokenExpiry(undefined, [
        { name: 'rtFa', expires: b },
        { name: 'FedAuth', expires: a },
      ]),
    ).toBe(new Date(a * 1000).toISOString());
  });
  it('uses the conservative window for session cookies', () => {
    const got = Date.parse(deriveTokenExpiry(undefined, [{ name: 'FedAuth', expires: -1 }]));
    expect(got).toBeGreaterThan(Date.now());
  });
  it('falls through to the cookie window when the bearer is malformed', () => {
    expect(() => deriveTokenExpiry('not-a-jwt', [{ name: 'FedAuth', expires: -1 }])).not.toThrow();
  });
});

describe('collectCookieHeader', () => {
  const all = [
    { name: 'FedAuth', value: 'a', domain: 'x.sharepoint.com' },
    { name: 'rtFa', value: 'b', domain: '.sharepoint.com' },
    { name: 'junk', value: 'c', domain: 'login.microsoftonline.com' },
  ];
  it('keeps host and parent-domain cookies', () => {
    const h = collectCookieHeader(all, 'x.sharepoint.com');
    expect(h).toContain('FedAuth=a');
    expect(h).toContain('rtFa=b');
  });
  it('drops cookies from unrelated domains', () => {
    expect(collectCookieHeader(all, 'x.sharepoint.com')).not.toContain('junk');
  });
  it('keeps parent-domain cookies for the -my host, which is why one session covers both', () => {
    expect(collectCookieHeader(all, 'x-my.sharepoint.com')).toContain('rtFa=b');
  });
});
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `npx vitest run test_scripts/lock.spec.ts test_scripts/capture.spec.ts`
Expected: FAIL.

- [ ] **Step 4: Implement `src/auth/capture.ts`**

Export the four pure helpers separately from the Playwright orchestration so they are testable without a browser. `captureSession` differs from the ported original in three ways:

1. It launches its own persistent context rather than receiving one.
2. Interactive mode waits for the user to complete sign-in: navigate, then poll `/_api/web` through the page context until it returns 200 or `timeoutMs` elapses, rather than assuming the first navigation lands authenticated.
3. Headless mode uses the short renew timeout and fails with `AUTH_REQUIRED` rather than hanging, so cron gets a clean signal to alert on.

Bearer capture stays best-effort via `context.on('request')` including `.mcas.ms`, and a missing bearer is not an error.

- [ ] **Step 5: Implement `login` and `auth-renew` commands**

Both acquire the lock, call `captureSession`, and `saveSession`. `login` uses `loginTimeoutMs` and `headless: false`; `auth-renew` uses `renewTimeoutMs` and `headless: true`.

- [ ] **Step 6: Run tests to verify they pass**

Run: `npm test`
Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "feat(auth): standalone capture, lock, login and auth-renew"
```

---

### Task 8: auth-check, health-check, output and CLI wiring

**Files:**

- Create: `src/commands/auth-check.ts`, `src/commands/health-check.ts`, `src/output/json.ts`
- Modify: `src/cli.ts` (replace the scaffold), `src/util/exit-codes.ts` (remove `NotImplemented`)
- Test: `test_scripts/auth-check.spec.ts`

**Interfaces:**

- Produces: `runAuthCheck(client): Promise<{ overall: 'ok' | 'degraded' | 'broken'; probes: Probe[] }>` where `Probe = { name: 'read' | 'write' | 'search'; ok: boolean; status?: number; detail: string; durationMs: number }`

- [ ] **Step 1: Write the failing test**

The point of this command is that it probes every dependency, so the tests assert exactly that.

```typescript
import { describe, expect, it, vi } from 'vitest';
import { runAuthCheck } from '../src/commands/auth-check';

const okClient = () => ({
  getJson: vi.fn().mockResolvedValue({ Title: 'T' }),
  postJson: vi.fn().mockResolvedValue({ FormDigestValue: 'D', FormDigestTimeoutSeconds: 1800 }),
});

describe('runAuthCheck', () => {
  it('probes read, write and search, not just one', async () => {
    const c = okClient();
    const r = await runAuthCheck(c as never);
    expect(r.probes.map((p) => p.name).sort()).toEqual(['read', 'search', 'write']);
    expect(r.overall).toBe('ok');
  });

  it('reports degraded when only the search probe fails', async () => {
    const c = okClient();
    c.getJson = vi
      .fn()
      .mockImplementation((u: string) =>
        u.includes('/_api/search')
          ? Promise.reject(new Error('no search'))
          : Promise.resolve({ Title: 'T' }),
      );
    const r = await runAuthCheck(c as never);
    expect(r.overall).toBe('degraded');
    expect(r.probes.find((p) => p.name === 'search')?.ok).toBe(false);
    expect(r.probes.find((p) => p.name === 'read')?.ok).toBe(true);
  });

  it('reports degraded when only the write probe fails, which the teams bug would have missed', async () => {
    const c = okClient();
    c.postJson = vi.fn().mockRejectedValue(new Error('no digest'));
    const r = await runAuthCheck(c as never);
    expect(r.overall).toBe('degraded');
    expect(r.probes.find((p) => p.name === 'write')?.ok).toBe(false);
  });

  it('reports broken when the read probe fails', async () => {
    const c = okClient();
    c.getJson = vi.fn().mockRejectedValue(new Error('401'));
    c.postJson = vi.fn().mockRejectedValue(new Error('401'));
    expect((await runAuthCheck(c as never)).overall).toBe('broken');
  });

  it('records a duration for every probe', async () => {
    const r = await runAuthCheck(okClient() as never);
    for (const p of r.probes) expect(typeof p.durationMs).toBe('number');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test_scripts/auth-check.spec.ts`
Expected: FAIL.

- [ ] **Step 3: Implement `auth-check`, `health-check` and `output/json.ts`**

`overall` is `broken` when the read probe fails, since nothing works without it, `degraded` when read passes but write or search fails, `ok` only when all three pass. `health-check` reuses `runAuthCheck` and adds nothing but formatting: one implementation, so the two can never disagree the way `teams-cli`'s do.

- [ ] **Step 4: Rewrite `src/cli.ts`**

Wire every command through commander with global `--host`, `--timeout`, `--session-file`, `--chrome-channel`. Every action catches `CliError` and `SharepointHttpError`, writes the JSON error payload to stderr, and exits with the mapped code. Remove `ExitCode.NotImplemented` and its test (resolves P5).

- [ ] **Step 5: Run the full suite and build**

Run: `npm test && npx tsc --noEmit && npm run lint && npm run build`
Expected: all green.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "feat(cli): auth-check with three probes, output layer and full wiring"
```

---

### Task 9: Live verification against the tenant

**Files:**

- Create: `test_scripts/live-smoke.sh`
- Modify: `docs/design/project-functions.MD` (states to `verified`), `Issues - Pending Items.md`

- [ ] **Step 1: Seed the dev session**

```bash
mkdir -p ~/.sharepoint-cli && chmod 700 ~/.sharepoint-cli
cp ~/.outlook-cli/sharepoint-session.json ~/.sharepoint-cli/session.json
chmod 600 ~/.sharepoint-cli/session.json
```

- [ ] **Step 2: Verify the read surface on both hosts**

```bash
H=$(python3 -c "import json,os;print(json.load(open(os.path.expanduser('~/.outlook-cli/sharepoint-session.json')))['host'])")
MY="${H%%.*}-my.${H#*.}"
node dist/cli.js --host "$H" auth-check
node dist/cli.js --host "$H" libraries
node dist/cli.js --host "$H" ls /
node dist/cli.js --host "$MY" auth-check
```

Expected: `auth-check` reports all three probes ok on both hosts; `libraries` lists the document libraries including the Greek-named ones; `ls /` succeeds.

- [ ] **Step 3: Verify the write surface in a scratch folder**

Ask the operator to nominate a scratch folder. Create a subfolder, upload a small Greek-named file, list it, download it, and compare bytes. Do not test chunked upload against the live tenant without explicit approval, since it writes tens of megabytes.

- [ ] **Step 4: Update the functional register and commit**

Mark verified functions `verified`, leaving `login` and `auth-renew` as `built` with a note that interactive verification is outstanding. Close resolved items in `Issues - Pending Items.md`.

```bash
git add -A
git commit -m "test: live verification of read and write surfaces on both hosts"
```

---

## Self-review

**Spec coverage.** Design §6 auth maps to Tasks 7 and 8; §7.1 paths to Task 2; §7.2 digest to Task 4; §7.3 upload to Task 6; §8 command surface to Tasks 5, 6 and 8; §9 error handling to Task 3; §12 testing throughout; §2.2 dual-host to Tasks 1, 3 and 9. Design §10 (MCP `files` plugin) is phase 4 and deliberately out of scope. §11 phase 5 (removal and the atm-recon migration) likewise.

**Placeholders.** None. Every code step carries real code; the three steps that describe rather than show (Task 5 step 3, Task 6 step 4, Task 8 step 3) give the exact URLs, the exact state-machine rules, and the exact test cases.

**Type consistency.** `SharepointSession` is defined once in Task 1 and consumed unchanged. `SharepointClient` methods are fixed in Task 3 and every later task uses those names. `DigestCache.get(force)` matches the `DigestProvider = (force: boolean) => Promise<string>` the client expects. `normalizeServerRelative`, `odataLiteral`, `folderApi`, `fileApi`, `splitParentLeaf` are defined in Task 2 and used with those names in Tasks 5 and 6.

**Known gap:** interactive `login` and headless `auth-renew` cannot be verified without a human at an MFA prompt. Task 9 leaves them `built`, not `verified`, and the final report must say so rather than implying end-to-end proof.
