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
  contextInfo<T>(): Promise<T>;
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
    private readonly client: DigestSource,
    private readonly now: () => number = () => Date.now(),
  ) {}

  async get(force = false): Promise<string> {
    if (!force && this.value !== null && this.now() < this.expiresAtMs) {
      return this.value;
    }
    // Collapse concurrent callers onto a single in-flight request.
    if (!force && this.inflight) return this.inflight;

    const p = this.fetch();
    this.inflight = p;
    try {
      return await p;
    } finally {
      // Clear on failure too, so a transient error is not cached forever.
      if (this.inflight === p) this.inflight = null;
    }
  }

  private async fetch(): Promise<string> {
    const info = await this.client.contextInfo<ContextInfo>();
    const digest = info?.FormDigestValue;
    if (typeof digest !== 'string' || digest.length === 0) {
      throw new Error('contextinfo returned no FormDigestValue');
    }
    const ttlS =
      typeof info.FormDigestTimeoutSeconds === 'number' && info.FormDigestTimeoutSeconds > 0
        ? info.FormDigestTimeoutSeconds
        : FALLBACK_TTL_S;
    this.value = digest;
    // With the fallback TTL this lands at `now`, so the next call refetches.
    // That pessimism is intended when SharePoint omits the field.
    this.expiresAtMs = this.now() + ttlS * 1000 - SAFETY_MARGIN_MS;
    return digest;
  }
}
