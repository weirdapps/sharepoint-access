// src/auth/capture.ts
//
// Playwright capture of a SharePoint session. Ported from
// outlook-access/src/auth/sharepoint-capture.ts and made standalone.
//
// The substantive change is COLD-PROFILE support. The original assumed a warm
// context whose Microsoft SSO cookies had just been set by an Outlook sign-in,
// so it navigated once and expected silence. Standalone, the first `login`
// meets the full interactive redirect chain including MFA, so interactive mode
// polls a cheap authenticated endpoint until it answers 200.
//
// A missing Bearer is NOT an error: cookie-auth (MCAS-gated) tenants emit
// none, and FedAuth/rtFa authorise on their own.

import type { BrowserContext, Request as PWRequest } from 'playwright';

import { CliError } from '../config/errors';
import type { SharepointSession } from '../session/schema';

export interface CaptureOptions {
  host: string;
  profileDir: string;
  chromeChannel: string;
  timeoutMs: number;
  /** true for silent renewal, false for interactive first login. */
  headless: boolean;
}

/** Cookie-auth sessions have no JWT to expire against. */
const COOKIE_FALLBACK_TTL_MS = 7 * 24 * 60 * 60 * 1000;
/** How often interactive mode re-checks whether sign-in has completed. */
const POLL_INTERVAL_MS = 2_000;

export interface ExpiringCookie {
  name: string;
  expires?: number;
}

export interface NamedCookie {
  name: string;
  value: string;
  domain: string;
}

// ── Pure helpers (unit-tested without a browser) ────────────────────────────

/**
 * Match a Bearer-carrying request to the SharePoint host, whether direct or
 * rewritten by MCAS (Defender for Cloud Apps), which proxies through a
 * "<original-fqdn>.mcas.ms" domain so the host no longer prefixes the URL.
 */
export function isSharepointBearerUrl(host: string, url: string): boolean {
  if (url.startsWith(`https://${host}/`)) return true;
  if (/^https:\/\/[^/]*\.mcas\.ms\//i.test(url)) {
    const tenant = host.split('.')[0].toLowerCase();
    // Check only the AUTHORITY, so a lookalike host cannot smuggle the tenant
    // name through as a path segment.
    const authority = url.slice('https://'.length).split('/')[0].toLowerCase();
    return authority.includes(tenant) && authority.includes('sharepoint');
  }
  return false;
}

function decodeJwtExp(jwt: string): number | null {
  const parts = jwt.split('.');
  if (parts.length !== 3) return null;
  try {
    const claims = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as {
      exp?: unknown;
    };
    return typeof claims.exp === 'number' ? claims.exp : null;
  } catch {
    return null;
  }
}

/**
 * JWT exp when a Bearer exists, else the FedAuth (then rtFa) cookie expiry,
 * else a conservative window. Session cookies report expires = -1.
 */
export function deriveTokenExpiry(
  bearer: string | undefined,
  cookies: ExpiringCookie[],
  now: () => number = () => Date.now(),
): string {
  if (bearer) {
    const exp = decodeJwtExp(bearer);
    if (exp !== null) return new Date(exp * 1000).toISOString();
  }
  for (const name of ['FedAuth', 'rtFa']) {
    const c = cookies.find((k) => k.name.toLowerCase() === name.toLowerCase());
    if (c && typeof c.expires === 'number' && c.expires > 0) {
      return new Date(c.expires * 1000).toISOString();
    }
  }
  return new Date(now() + COOKIE_FALLBACK_TTL_MS).toISOString();
}

/**
 * Serialize the cookies for `host` AND its parent domain. The parent-domain
 * part is why one session covers both the team-sites host and the -my
 * OneDrive for Business host.
 */
export function collectCookieHeader(all: NamedCookie[], host: string): string {
  const parent = host.split('.').slice(-2).join('.');
  return all
    .filter(
      (c) =>
        c.domain === host ||
        c.domain === `.${host}` ||
        c.domain === parent ||
        c.domain === `.${parent}`,
    )
    .map((c) => `${c.name}=${c.value}`)
    .join('; ');
}

// ── Playwright orchestration ────────────────────────────────────────────────

/** Probe an authenticated endpoint from inside the browser context. */
async function isSignedIn(context: BrowserContext, host: string): Promise<boolean> {
  try {
    const resp = await context.request.get(`https://${host}/_api/web?$select=Title`, {
      headers: { Accept: 'application/json;odata=nometadata' },
      timeout: 15_000,
    });
    return resp.status() === 200;
  } catch {
    return false;
  }
}

export async function captureSession(opts: CaptureOptions): Promise<SharepointSession> {
  const { chromium } = await import('playwright');

  const fs = await import('node:fs');
  fs.mkdirSync(opts.profileDir, { recursive: true, mode: 0o700 });

  const context = await chromium.launchPersistentContext(opts.profileDir, {
    channel: opts.chromeChannel,
    headless: opts.headless,
    viewport: { width: 1280, height: 900 },
    args: ['--no-first-run', '--no-default-browser-check'],
  });

  let capturedAuth: string | null = null;
  const onRequest = (req: PWRequest): void => {
    if (capturedAuth) return;
    try {
      if (!isSharepointBearerUrl(opts.host, req.url())) return;
      const header = req.headers()['authorization'] ?? '';
      if (/^Bearer\s+/i.test(header)) capturedAuth = header;
    } catch {
      /* best-effort: a malformed request must not abort capture */
    }
  };
  context.on('request', onRequest);

  try {
    const page = await context.newPage();
    try {
      await page.goto(`https://${opts.host}/_layouts/15/sharepoint.aspx`, {
        waitUntil: 'domcontentloaded',
        timeout: Math.min(opts.timeoutMs, 60_000),
      });
    } catch {
      // Navigation can error under MCAS redirects; cookies may still be set,
      // and the poll below is the real completion test.
    }

    // Cold profile: wait for the human to finish signing in. Warm profile:
    // the first poll succeeds immediately.
    const deadline = Date.now() + opts.timeoutMs;
    let signedIn = await isSignedIn(context, opts.host);
    while (!signedIn && Date.now() < deadline) {
      if (opts.headless) break; // no human to wait for
      await page.waitForTimeout(POLL_INTERVAL_MS);
      signedIn = await isSignedIn(context, opts.host);
    }

    if (!signedIn) {
      throw new CliError(
        'AUTH_REQUIRED',
        opts.headless
          ? `silent renewal failed for ${opts.host}: run "sharepoint-cli login --host ${opts.host}"`
          : `sign-in did not complete within ${opts.timeoutMs}ms for ${opts.host}`,
      );
    }

    const bearer = capturedAuth ? (capturedAuth as string).replace(/^Bearer\s+/i, '') : undefined;
    const all = await context.cookies();
    const cookies = collectCookieHeader(all, opts.host);

    if (!bearer && !cookies) {
      throw new CliError(
        'AUTH_REQUIRED',
        `no SharePoint auth captured for ${opts.host}: no Bearer and no cookies`,
      );
    }

    return {
      version: 1,
      host: opts.host,
      ...(bearer ? { bearer } : {}),
      cookies,
      capturedAt: new Date().toISOString(),
      tokenExpiresAt: deriveTokenExpiry(
        bearer,
        all.filter((c) => c.domain.includes('sharepoint.com')),
      ),
    };
  } finally {
    context.off('request', onRequest);
    await context.close().catch(() => {
      /* tolerate teardown races */
    });
  }
}
