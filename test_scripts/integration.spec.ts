// test_scripts/integration.spec.ts
//
// Wires the REAL SharepointClient to the REAL DigestCache, mocking only
// `fetch`. Everything else in the suite mocks at the client boundary, which is
// exactly why the digest recursion bug survived 214 unit tests and only showed
// up against the live tenant: postJson asked the digest provider for a header,
// and the provider's only way to get one was postJson.

import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';

import { SharepointClient } from '../src/http/client';
import { DigestCache } from '../src/sharepoint/digest';
import { runMkdir } from '../src/commands/mkdir';
import { runAuthCheck } from '../src/commands/auth-check';

const session = {
  version: 1 as const,
  host: 'x.sharepoint.com',
  cookies: 'FedAuth=aaa',
  capturedAt: '2026-08-08T00:00:00.000Z',
  tokenExpiresAt: '2099-01-01T00:00:00.000Z',
};

/** Build the same object graph cli.ts builds. */
function wired() {
  const client = new SharepointClient(session, { httpTimeoutMs: 5000 });
  const digest = new DigestCache(client);
  client.setDigestProvider((force) => digest.get(force));
  return client;
}

const CONTEXTINFO = JSON.stringify({
  FormDigestValue: 'DIGEST-1',
  FormDigestTimeoutSeconds: 1800,
});

describe('client wired to digest cache', () => {
  let fetchMock: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    fetchMock = vi
      .fn()
      .mockImplementation((url: string) =>
        Promise.resolve(
          url.endsWith('/_api/contextinfo')
            ? new Response(CONTEXTINFO, { status: 200 })
            : new Response('{"ok":true}', { status: 200 }),
        ),
      );
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('does not recurse when a write needs a digest', async () => {
    await expect(runMkdir(wired(), '/a/New')).resolves.toBeDefined();
  });

  it('fetches contextinfo exactly once, then the write', async () => {
    await runMkdir(wired(), '/a/New');
    const urls = fetchMock.mock.calls.map((c) => c[0] as string);
    expect(urls.filter((u) => u.endsWith('/_api/contextinfo')).length).toBe(1);
    expect(urls.some((u) => u.includes('addUsingPath'))).toBe(true);
  });

  it('sends the digest it just minted on the write request', async () => {
    await runMkdir(wired(), '/a/New');
    const write = fetchMock.mock.calls.find((c) => (c[0] as string).includes('addUsingPath'));
    expect((write?.[1].headers as Record<string, string>)['X-RequestDigest']).toBe('DIGEST-1');
  });

  it('does not attach a digest header to the contextinfo request itself', async () => {
    await runMkdir(wired(), '/a/New');
    const ctx = fetchMock.mock.calls.find((c) => (c[0] as string).endsWith('/_api/contextinfo'));
    expect((ctx?.[1].headers as Record<string, string>)['X-RequestDigest']).toBeUndefined();
  });

  it('reuses the cached digest across two writes', async () => {
    const client = wired();
    await runMkdir(client, '/a/One');
    await runMkdir(client, '/a/Two');
    const urls = fetchMock.mock.calls.map((c) => c[0] as string);
    expect(urls.filter((u) => u.endsWith('/_api/contextinfo')).length).toBe(1);
  });

  it('mints the digest from the sub-site web when writing to a sub-site', async () => {
    // Live-verified 2026-08-08: a root-web digest presented to a /personal
    // web is rejected 403 with the stale-validation marker. Getting this wrong
    // breaks every write outside the root web, and the error reads like
    // expiry rather than mis-scoping.
    await runMkdir(wired(), '/personal/u/Documents/New');
    const ctx = fetchMock.mock.calls.find((c) => (c[0] as string).endsWith('/_api/contextinfo'));
    expect(ctx?.[0]).toBe('https://x.sharepoint.com/personal/u/_api/contextinfo');
    const write = fetchMock.mock.calls.find((c) => (c[0] as string).includes('addUsingPath'));
    expect(write?.[0]).toContain('/personal/u/_api/web/folders/');
  });

  it('keeps a separate cached digest per web', async () => {
    const client = wired();
    await runMkdir(client, '/personal/u/Documents/A');
    await runMkdir(client, '/sites/s/Docs/B');
    const ctxUrls = fetchMock.mock.calls
      .map((c) => c[0] as string)
      .filter((u) => u.endsWith('/_api/contextinfo'));
    expect(ctxUrls).toEqual([
      'https://x.sharepoint.com/personal/u/_api/contextinfo',
      'https://x.sharepoint.com/sites/s/_api/contextinfo',
    ]);
  });

  it('runs a full auth-check against the wired graph without recursing', async () => {
    const r = await runAuthCheck(wired());
    expect(r.overall).toBe('ok');
    expect(r.probes.find((p) => p.name === 'write')?.ok).toBe(true);
  });

  it('reports the write probe as failed, not as a stack overflow, when contextinfo 401s', async () => {
    fetchMock.mockImplementation((url: string) =>
      Promise.resolve(
        url.endsWith('/_api/contextinfo')
          ? new Response('denied', { status: 401 })
          : new Response('{"ok":true}', { status: 200 }),
      ),
    );
    const r = await runAuthCheck(wired());
    expect(r.overall).toBe('degraded');
    const write = r.probes.find((p) => p.name === 'write');
    expect(write?.ok).toBe(false);
    expect(write?.detail).not.toMatch(/call stack/i);
  });
});
