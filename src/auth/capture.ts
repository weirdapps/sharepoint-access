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
// The cookies ARE the session. An earlier version also scavenged a Bearer out
// of the page's own traffic and stored it; the client never sent it, and it
// aged out inside the hour while FedAuth/rtFa stayed valid for days, so the
// only thing it contributed was a live credential sitting unused on disk.

import type { BrowserContext } from 'playwright';

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
 * The FedAuth (then rtFa) cookie expiry, else a conservative window. Session
 * cookies report expires = -1.
 *
 * This used to prefer the Bearer's JWT exp, and that made the field describe a
 * credential the client does not send. The session captured on 2026-09-10 at
 * 16:54:54Z advertised 14:11:39Z, sixteen minutes and forty-five seconds, while
 * the FedAuth cookie beside it was valid until 09-15. Every consumer reading
 * tokenExpiresAt was told a five-day session had a quarter of an hour left.
 */
export function deriveTokenExpiry(
  cookies: ExpiringCookie[],
  now: () => number = () => Date.now(),
): string {
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

  // No Bearer is scavenged from the page's traffic any more. The client sends
  // cookies only, so a captured Bearer was an unused credential written to disk
  // on three machines and copied between them every fifteen minutes.
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

    const all = await context.cookies();
    const cookies = collectCookieHeader(all, opts.host);

    if (!cookies) {
      throw new CliError(
        'AUTH_REQUIRED',
        `no SharePoint auth captured for ${opts.host}: no cookies`,
      );
    }

    return {
      version: 1,
      host: opts.host,
      cookies,
      capturedAt: new Date().toISOString(),
      tokenExpiresAt: deriveTokenExpiry(all.filter((c) => c.domain.includes('sharepoint.com'))),
    };
  } finally {
    await context.close().catch(() => {
      /* tolerate teardown races */
    });
  }
}
