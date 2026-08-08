// src/http/client.ts
//
// Cookie-authenticated classic SPO REST client.
//
// Two security-relevant rules live here:
//   1. An absolute URL is host-checked BEFORE the session cookies are
//      attached. Without that, a crafted attachment or share URL would
//      exfiltrate the session to an arbitrary host.
//   2. Error text never interpolates the session. Callers log messages.

import { Agent } from 'undici';

import type { SharepointSession } from '../session/schema';
import { CliError } from '../config/errors';
import { SharepointHttpError, isStaleDigest } from './errors';

export interface ClientOpts {
  httpTimeoutMs: number;
  /** TCP connect ceiling. Separate from the overall request timeout. */
  connectTimeoutMs?: number;
}

/**
 * undici's DEFAULT CONNECT TIMEOUT IS 10s AND IS NOT CONFIGURABLE WITHOUT A
 * DISPATCHER. Measured against the tenant on 2026-08-08: some SharePoint
 * front-ends take anywhere from 1.5s to 42s just to complete the TCP connect,
 * so with the default roughly two requests in three died with
 * UND_ERR_CONNECT_TIMEOUT while the identical request from Python succeeded.
 * A dispatcher with a generous connect timeout is the fix; the retry below
 * only covers the residue.
 */
const DEFAULT_CONNECT_TIMEOUT_MS = 30_000;

/** Digests are web-scoped, so the provider is asked for a specific web. */
export type DigestProvider = (web: string, force: boolean) => Promise<string>;

/**
 * The web a request targets, i.e. everything before "/_api/".
 * "/personal/u/_api/web/folders/..." -> "/personal/u"
 * "/_api/web/lists"                  -> ""
 */
export function webOfPath(path: string): string {
  const i = path.indexOf('/_api/');
  return i <= 0 ? '' : path.slice(0, i);
}

export interface BinaryResult {
  bytes: Buffer;
  contentType: string;
  filename?: string;
}

/** Extra attempts after the first, for connection-level failures only. */
const NETWORK_RETRIES = 2;

/** undici error codes that mean "never reached the server", so a retry is safe. */
const RETRYABLE_CODES = new Set([
  'UND_ERR_CONNECT_TIMEOUT',
  'UND_ERR_SOCKET',
  'ECONNRESET',
  'ECONNREFUSED',
  'ETIMEDOUT',
  'EAI_AGAIN',
  'ENETUNREACH',
  'EPIPE',
]);

function errorCodeOf(err: unknown): string | undefined {
  const cause = (err as { cause?: unknown })?.cause;
  return (
    (err as { code?: string })?.code ?? (cause as { code?: string } | undefined)?.code ?? undefined
  );
}

export function isRetryableNetworkError(err: unknown): boolean {
  const code = errorCodeOf(err);
  return code !== undefined && RETRYABLE_CODES.has(code);
}

/**
 * Node's fetch throws a bare TypeError("fetch failed") and hides the real
 * reason on `.cause`. Unwrap it so operators see a diagnosable message.
 */
export function describeFetchError(err: unknown): string {
  const msg = (err as Error)?.message ?? String(err);
  const cause = (err as { cause?: unknown })?.cause;
  if (!cause) return msg;
  const causeMsg = (cause as Error)?.message ?? String(cause);
  const code = errorCodeOf(err);
  return code ? `${msg}: ${causeMsg} (${code})` : `${msg}: ${causeMsg}`;
}

/** Only https *.sharepoint.com absolute URLs may receive the session cookies. */
export function assertSharepointUrl(raw: string): URL {
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
  // Note the leading dot: "evilsharepoint.com" must not pass a suffix test.
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

  private dispatcher: Agent | undefined;

  constructor(
    private readonly session: SharepointSession,
    private readonly opts: ClientOpts,
  ) {
    try {
      this.dispatcher = new Agent({
        connect: { timeout: opts.connectTimeoutMs ?? DEFAULT_CONNECT_TIMEOUT_MS },
      });
    } catch {
      this.dispatcher = undefined;
    }
  }

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
    if (/^[a-z][a-z0-9+.-]*:/i.test(pathOrUrl)) return assertSharepointUrl(pathOrUrl).toString();
    // A relative path MUST start with a single "/". Without the leading slash
    // "evil.com/x" would concatenate into "https://<host>evil.com/x", pointing
    // at a different host with our cookies attached. Without the not-"//"
    // check, "//evil.com/x" would become a protocol-relative authority. No
    // current caller can reach either, since every path is built by folderApi
    // or fileApi; this stops a future one from reintroducing the hole.
    if (!pathOrUrl.startsWith('/') || pathOrUrl.startsWith('//')) {
      throw new CliError(
        'CONFIG_INVALID',
        `internal: request path must start with a single "/", got "${pathOrUrl}"`,
      );
    }
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
    let lastErr: unknown;

    // Retry transient CONNECTION failures only. Measured against the tenant on
    // 2026-08-08: some SharePoint front-ends take ~10s to complete a TCP
    // connect, right at undici's fixed 10s connect timeout, so identical
    // requests fail roughly two times in three and then succeed. undici's
    // connect timeout is not adjustable without supplying a dispatcher, and a
    // bounded retry is the smaller change. Only pre-response errors are
    // retried: once the server has answered, the status is the answer.
    for (let attempt = 0; attempt <= NETWORK_RETRIES; attempt++) {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), this.opts.httpTimeoutMs);
      // Capture what THIS attempt used. Concurrent requests share the client,
      // so a sibling can clear the dispatcher between here and the catch;
      // testing `this.dispatcher` there would make every sibling skip its own
      // retry and fail. auth-check runs three probes in parallel and hit
      // exactly that.
      const usedDispatcher = this.dispatcher !== undefined;
      try {
        return await fetch(url, {
          method,
          headers: { ...this.baseHeaders(accept), ...(extra ?? {}) },
          body,
          signal: ctrl.signal,
          redirect: 'follow',
          ...(this.dispatcher ? { dispatcher: this.dispatcher } : {}),
        } as RequestInit & { dispatcher?: Agent });
      } catch (err) {
        lastErr = err;
        if ((err as Error).name === 'AbortError') {
          throw new CliError(
            'TIMEOUT',
            `request timed out after ${this.opts.httpTimeoutMs}ms: ${url}`,
          );
        }
        // The standalone undici Agent is only accepted by a built-in fetch of a
        // matching undici version. Node 26 accepts it; Node 24 on the VPS
        // rejects it with UND_ERR_INVALID_ARG before a single byte is sent.
        // Drop the dispatcher for the rest of the process and retry: a longer
        // connect timeout is an optimisation, never a requirement, and the
        // alternative is a CLI that simply does not run on the VPS.
        if (errorCodeOf(err) === 'UND_ERR_INVALID_ARG' && usedDispatcher) {
          this.dispatcher = undefined;
          attempt--; // this attempt never reached the network
          continue;
        }
        // A body cannot be replayed once consumed, so never retry those.
        if (!isRetryableNetworkError(err) || body !== undefined) break;
      } finally {
        clearTimeout(timer);
      }
    }

    // Surface the underlying cause. Bare "fetch failed" is undiagnosable, and
    // cost real time when this endpoint first misbehaved.
    throw new CliError('UPSTREAM', `request failed: ${describeFetchError(lastErr)}`);
  }

  async getJson<T>(path: string): Promise<T> {
    const resp = await this.raw('GET', path, 'application/json;odata=nometadata');
    if (!resp.ok) throw new SharepointHttpError(resp.status, this.url(path), await resp.text());
    return (await resp.json()) as T;
  }

  async getBinary(pathOrUrl: string): Promise<BinaryResult> {
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
   * envelope in outlook-access/src/http/outlook-client.ts:435, so all three
   * CLIs behave alike under credential churn.
   */
  private async post(
    path: string,
    body: BodyInit | undefined,
    extra: Record<string, string>,
    accept: string,
  ): Promise<Response> {
    // The digest must come from the SAME web the write targets: one minted at
    // the host root is rejected 403 by a sub-site web, with the stale-digest
    // marker, which reads like expiry rather than mis-scoping.
    const web = webOfPath(path);
    const send = async (force: boolean): Promise<Response> => {
      const headers = { ...extra };
      if (this.digestProvider) headers['X-RequestDigest'] = await this.digestProvider(web, force);
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

  /**
   * POST /_api/contextinfo WITHOUT a digest header.
   *
   * This must never go through `post()`. That path asks the digest provider
   * for a digest, and the digest provider's only way to get one is this call:
   * routing it through `post()` recurses until the stack blows. Deliberately a
   * distinct method rather than a flag, so the cycle cannot be reintroduced by
   * passing the wrong argument.
   */
  async contextInfo<T>(web = ''): Promise<T> {
    const path = `${web}/_api/contextinfo`;
    const resp = await this.raw('POST', path, 'application/json;odata=nometadata');
    if (!resp.ok) {
      throw new SharepointHttpError(resp.status, this.url(path), await resp.text());
    }
    const text = await resp.text();
    return (text.length > 0 ? JSON.parse(text) : {}) as T;
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
  // RFC 5987 filename*=UTF-8''… takes precedence over the plain form.
  const m1 = header.match(/filename\*=(?:UTF-8'')?([^;]+)/i);
  if (m1) {
    try {
      return decodeURIComponent(m1[1].trim().replace(/^"|"$/g, ''));
    } catch {
      /* malformed encoding: fall through to the plain form */
    }
  }
  const m2 = header.match(/filename=("?)([^";]+)\1/i);
  return m2 ? m2[2].trim() : undefined;
}
