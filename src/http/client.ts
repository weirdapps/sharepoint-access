// src/http/client.ts
//
// Cookie-authenticated classic SPO REST client.
//
// Two security-relevant rules live here:
//   1. An absolute URL is host-checked BEFORE the session cookies are
//      attached. Without that, a crafted attachment or share URL would
//      exfiltrate the session to an arbitrary host.
//   2. Error text never interpolates the session. Callers log messages.

import type { SharepointSession } from '../session/schema';
import { CliError } from '../config/errors';
import { SharepointHttpError, isStaleDigest } from './errors';

export interface ClientOpts {
  httpTimeoutMs: number;
}

export type DigestProvider = (force: boolean) => Promise<string>;

export interface BinaryResult {
  bytes: Buffer;
  contentType: string;
  filename?: string;
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

  /**
   * POST /_api/contextinfo WITHOUT a digest header.
   *
   * This must never go through `post()`. That path asks the digest provider
   * for a digest, and the digest provider's only way to get one is this call:
   * routing it through `post()` recurses until the stack blows. Deliberately a
   * distinct method rather than a flag, so the cycle cannot be reintroduced by
   * passing the wrong argument.
   */
  async contextInfo<T>(): Promise<T> {
    const resp = await this.raw('POST', '/_api/contextinfo', 'application/json;odata=nometadata');
    if (!resp.ok) {
      throw new SharepointHttpError(resp.status, this.url('/_api/contextinfo'), await resp.text());
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
