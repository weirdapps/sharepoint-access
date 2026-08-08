// test_scripts/security.spec.ts
//
// Regression tests for findings from the 2026-08-08 security audit. Each one
// fails if the corresponding defence is removed.

import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';

import { runSearch } from '../src/commands/search';
import { SharepointClient, assertSharepointUrl } from '../src/http/client';
import { normalizeServerRelative, encodeLeaf } from '../src/sharepoint/paths';

const session = {
  version: 1 as const,
  host: 'x.sharepoint.com',
  cookies: 'FedAuth=SECRET',
  capturedAt: '2026-08-08T00:00:00.000Z',
  tokenExpiresAt: '2099-01-01T00:00:00.000Z',
};

describe('AUDIT-1: prototype pollution via search result cell keys', () => {
  const rowWith = (cells: Array<{ Key: string; Value: string }>) => ({
    PrimaryQueryResult: { RelevantResults: { Table: { Rows: [{ Cells: cells }] } } },
  });

  it('does not pollute Object.prototype via a __proto__ cell key', async () => {
    const getJson = vi.fn().mockResolvedValue(rowWith([{ Key: '__proto__', Value: 'polluted' }]));
    await runSearch({ getJson } as never, 'q', 1);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    expect(Object.prototype).not.toHaveProperty('polluted');
  });

  it('drops __proto__, constructor and prototype keys from the result', async () => {
    const getJson = vi.fn().mockResolvedValue(
      rowWith([
        { Key: '__proto__', Value: 'a' },
        { Key: 'constructor', Value: 'b' },
        { Key: 'prototype', Value: 'c' },
        { Key: 'Title', Value: 'kept' },
      ]),
    );
    const r = await runSearch({ getJson } as never, 'q', 1);
    expect(r.results[0]).toEqual({ Title: 'kept' });
  });

  it('still returns a plain serialisable object', async () => {
    const getJson = vi.fn().mockResolvedValue(rowWith([{ Key: 'Title', Value: 'T' }]));
    const r = await runSearch({ getJson } as never, 'q', 1);
    expect(JSON.parse(JSON.stringify(r.results[0]))).toEqual({ Title: 'T' });
  });
});

describe('AUDIT-2: relative request paths cannot retarget the host', () => {
  let fetchMock: ReturnType<typeof vi.fn>;
  beforeEach(() => {
    fetchMock = vi.fn().mockResolvedValue(new Response('{}', { status: 200 }));
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('rejects a path with no leading slash, which would concatenate into another host', async () => {
    // "evil.com/x" would otherwise become "https://x.sharepoint.comevil.com/x".
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await expect(c.getJson('evil.com/x')).rejects.toThrowError(/start with a single/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('rejects a protocol-relative path', async () => {
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await expect(c.getJson('//evil.com/x')).rejects.toThrowError(/start with a single/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('accepts an ordinary API path', async () => {
    const c = new SharepointClient(session, { httpTimeoutMs: 1000 });
    await c.getJson('/_api/web');
    expect(fetchMock.mock.calls[0][0]).toBe('https://x.sharepoint.com/_api/web');
  });
});

describe('AUDIT-3: session cookies never leave *.sharepoint.com', () => {
  it.each([
    'https://evil.example.com/a',
    'https://evilsharepoint.com/a',
    'https://x.sharepoint.com.evil.com/a',
    'http://x.sharepoint.com/a',
    'file:///etc/passwd',
  ])('refuses %s', (url) => {
    expect(() => assertSharepointUrl(url)).toThrowError();
  });

  it.each(['https://x.sharepoint.com/a', 'https://x-my.sharepoint.com/a'])('allows %s', (url) => {
    expect(() => assertSharepointUrl(url)).not.toThrow();
  });
});

describe('AUDIT-4: traversal cannot escape the target folder', () => {
  it.each(['/a/../../etc/passwd', '/a//../b', '..', '/../x'])('rejects path %s', (p) => {
    expect(() => normalizeServerRelative(p)).toThrowError();
  });

  it.each(['../evil.docx', 'a/b.docx', '..', '.'])('rejects leaf %s', (leaf) => {
    expect(() => encodeLeaf(leaf)).toThrowError();
  });

  it('allows a Greek leaf with an apostrophe', () => {
    expect(() => encodeLeaf("Έκθεση O'B.docx")).not.toThrow();
  });
});

describe('AUDIT-5: credentials stay out of error text', () => {
  let fetchMock: ReturnType<typeof vi.fn>;
  beforeEach(() => {
    fetchMock = vi.fn();
    vi.stubGlobal('fetch', fetchMock);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('omits cookies and bearer from a thrown HTTP error', async () => {
    fetchMock.mockResolvedValue(new Response('body', { status: 500 }));
    const c = new SharepointClient({ ...session, bearer: 'SECRETJWT' }, { httpTimeoutMs: 1000 });
    const err = (await c.getJson('/_api/web').catch((e) => e)) as Error;
    const text = `${err.message}\n${err.stack ?? ''}`;
    expect(text).not.toContain('SECRETJWT');
    expect(text).not.toContain('SECRET');
  });
});
