// src/sharepoint/digest.ts
//
// X-RequestDigest cache. Process-local and never persisted: it is a CSRF
// token whose lifetime (1800s on this tenant) is shorter than most cron gaps,
// so writing it to disk would add attack surface for no benefit.

interface ContextInfo {
  FormDigestValue?: string;
  FormDigestTimeoutSeconds?: number;
}

/**
 * Structural type: keeps the cache testable without a real client.
 *
 * Deliberately requires `contextInfo`, NOT `postJson`. The digest-carrying
 * POST path calls this cache to obtain its header, so if the cache called
 * `postJson` back it would recurse until the stack blew. Naming a distinct
 * method makes that cycle unrepresentable.
 */
interface DigestSource {
  contextInfo<T>(web: string): Promise<T>;
}

/** Applied when SharePoint omits the timeout. Deliberately pessimistic. */
const FALLBACK_TTL_S = 60;
/** Covers clock skew and uploads that start just before expiry. */
const SAFETY_MARGIN_MS = 60_000;

interface Entry {
  value: string | null;
  expiresAtMs: number;
  inflight: Promise<string> | null;
}

/**
 * Keyed by WEB, because digests are web-scoped. Verified on the live tenant
 * 2026-08-08: a digest minted at the host root and presented to a
 * /personal/<user> web is rejected 403 with the stale-validation marker, while
 * one minted at that web succeeds. A single global cache therefore breaks
 * every write outside the root web, and does so with an error that reads like
 * expiry rather than mis-scoping.
 */
export class DigestCache {
  private readonly byWeb = new Map<string, Entry>();

  constructor(
    private readonly client: DigestSource,
    private readonly now: () => number = () => Date.now(),
  ) {}

  private entry(web: string): Entry {
    let e = this.byWeb.get(web);
    if (!e) {
      e = { value: null, expiresAtMs: 0, inflight: null };
      this.byWeb.set(web, e);
    }
    return e;
  }

  async get(web = '', force = false): Promise<string> {
    const e = this.entry(web);
    if (!force && e.value !== null && this.now() < e.expiresAtMs) {
      return e.value;
    }
    // Collapse concurrent callers for the SAME web onto one request.
    if (!force && e.inflight) return e.inflight;

    const p = this.fetch(web, e);
    e.inflight = p;
    try {
      return await p;
    } finally {
      // Clear on failure too, so a transient error is not cached forever.
      if (e.inflight === p) e.inflight = null;
    }
  }

  private async fetch(web: string, e: Entry): Promise<string> {
    const info = await this.client.contextInfo<ContextInfo>(web);
    const digest = info?.FormDigestValue;
    if (typeof digest !== 'string' || digest.length === 0) {
      throw new Error('contextinfo returned no FormDigestValue');
    }
    const ttlS =
      typeof info.FormDigestTimeoutSeconds === 'number' && info.FormDigestTimeoutSeconds > 0
        ? info.FormDigestTimeoutSeconds
        : FALLBACK_TTL_S;
    e.value = digest;
    // With the fallback TTL this lands at `now`, so the next call refetches.
    // That pessimism is intended when SharePoint omits the field.
    e.expiresAtMs = this.now() + ttlS * 1000 - SAFETY_MARGIN_MS;
    return digest;
  }
}
