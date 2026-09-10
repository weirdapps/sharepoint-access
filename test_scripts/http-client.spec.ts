import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';

import { SharepointClient } from '../src/http/client';
import { SharepointHttpError } from '../src/http/errors';

const STALE =
  '{"error":{"message":"The security validation for this page is invalid and might be corrupted."}}';

const session = {
  version: 1 as const,
  host: 'x.sharepoint.com',
  cookies: 'FedAuth=aaa',
  capturedAt: '2026-08-08T00:00:00.000Z',
  tokenExpiresAt: '2099-01-01T00:00:00.000Z',
};

const ok = (body: unknown) =>
  new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });

describe('SharepointClient', () => {
  let fetchMock: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    fetchMock = vi.fn();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  const headersOf = (i: number) => fetchMock.mock.calls[i][1].headers as Record<string, string>;

  it('sends the Cookie header and no Authorization when there is no bearer', async () => {
    fetchMock.mockResolvedValue(ok({ Title: 'T' }));
    await new SharepointClient(session, { httpTimeoutMs: 1000 }).getJson('/_api/web');
    expect(headersOf(0).Cookie).toBe('FedAuth=aaa');
    expect(headersOf(0).Authorization).toBeUndefined();
  });

  // The cookie is the credential; the bearer is not. Capture scavenges whatever
  // Bearer happens to be in flight, so it arrives with only the remainder of its
  // ~1h life -- 16m45s in the session that broke the VPS on 2026-09-10. The
  // FedAuth cookie beside it was good for five days. Sending the dead bearer
  // alongside the live cookie makes SharePoint answer 500 to every _api call:
  // measured on the VPS, the identical request with the Authorization header
  // removed returned 200. So the header is never an asset and eventually a
  // liability, and the fix is not to send it at all.
  it('never sends Authorization, even when the session carries a bearer', async () => {
    fetchMock.mockResolvedValue(ok({}));
    await new SharepointClient({ ...session, bearer: 'jwt' }, { httpTimeoutMs: 1000 }).getJson(
      '/_api/web',
    );
    expect(headersOf(0).Authorization).toBeUndefined();
    expect(headersOf(0).Cookie).toBe('FedAuth=aaa');
  });

  it('builds the URL from the session host', async () => {
    fetchMock.mockResolvedValue(ok({}));
    await new SharepointClient(session, { httpTimeoutMs: 1000 }).getJson('/_api/web');
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
    // (web, force): the write targets the root web here.
    expect(digest).toHaveBeenNthCalledWith(1, '', false);
    expect(digest).toHaveBeenNthCalledWith(2, '', true);
    expect(headersOf(1)['X-RequestDigest']).toBe('D2');
  });

  it('does not retry twice on a repeated stale-digest 403', async () => {
    // A fresh Response per call: a body can only be read once, and real fetch
    // never hands back the same object twice.
    fetchMock.mockImplementation(() => Promise.resolve(new Response(STALE, { status: 403 })));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    c.setDigestProvider(vi.fn().mockResolvedValue('D'));
    await expect(c.postJson('/_api/web/x')).rejects.toThrowError(SharepointHttpError);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('asks for a digest scoped to the sub-site web the write targets', async () => {
    fetchMock.mockResolvedValue(ok({}));
    const digest = vi.fn().mockResolvedValue('D');
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    c.setDigestProvider(digest);
    await c.postJson('/personal/u/_api/web/folders/addUsingPath');
    expect(digest).toHaveBeenCalledWith('/personal/u', false);
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

  it('rejects a lookalike host that merely ends in the string sharepoint.com', async () => {
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await expect(c.getBinary('https://evilsharepoint.com/a')).rejects.toThrowError(/host/i);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('rejects a non-https absolute URL', async () => {
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await expect(c.getBinary('http://x.sharepoint.com/a')).rejects.toThrowError(/https/i);
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
    const err = (await c.getJson('/_api/web').catch((e) => e)) as Error;
    expect(err.message).not.toContain('SECRETJWT');
    expect(err.message).not.toContain('FedAuth');
    expect(err.message).not.toContain('aaa');
  });

  it('maps an aborted request to TIMEOUT', async () => {
    fetchMock.mockImplementation(() => {
      const e = new Error('aborted');
      e.name = 'AbortError';
      return Promise.reject(e);
    });
    const c = new SharepointClient(session, { httpTimeoutMs: 5 });
    await expect(c.getJson('/_api/web')).rejects.toMatchObject({ code: 'TIMEOUT' });
  });

  it('parses an empty POST response body as an empty object', async () => {
    fetchMock.mockResolvedValue(new Response('', { status: 200 }));
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    expect(await c.postJson('/_api/web/x')).toEqual({});
  });

  it('extracts a UTF-8 filename from Content-Disposition', async () => {
    fetchMock.mockResolvedValue(
      new Response('bytes', {
        status: 200,
        headers: { 'content-disposition': "attachment; filename*=UTF-8''%CE%B1.docx" },
      }),
    );
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    expect((await c.getBinary('/_api/x')).filename).toBe('α.docx');
  });
});
