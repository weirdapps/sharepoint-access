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

  it('ignores a lookalike host carrying the tenant only as a path segment', () => {
    expect(isSharepointBearerUrl(host, 'https://evil.example.com/x.sharepoint.com/_api')).toBe(
      false,
    );
  });

  it('ignores an mcas.ms URL whose tenant marker is only in the path', () => {
    // The authority check exists precisely so this cannot pass.
    expect(isSharepointBearerUrl(host, 'https://other.mcas.ms/x-sharepoint/_api')).toBe(false);
  });

  it('does not match a host that merely has ours as a prefix', () => {
    expect(isSharepointBearerUrl(host, 'https://x.sharepoint.com.evil.com/_api')).toBe(false);
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

  it('matches the cookie name case-insensitively', () => {
    const secs = Math.floor(Date.UTC(2029, 0, 1) / 1000);
    expect(deriveTokenExpiry(undefined, [{ name: 'fedauth', expires: secs }])).toBe(
      new Date(secs * 1000).toISOString(),
    );
  });

  it('uses the conservative window for session cookies', () => {
    const now = Date.UTC(2026, 0, 1);
    const got = Date.parse(
      deriveTokenExpiry(undefined, [{ name: 'FedAuth', expires: -1 }], () => now),
    );
    expect(got).toBe(now + 7 * 24 * 60 * 60 * 1000);
  });

  it('falls through to the cookie window when the bearer is malformed', () => {
    const now = Date.UTC(2026, 0, 1);
    expect(() =>
      deriveTokenExpiry('not-a-jwt', [{ name: 'FedAuth', expires: -1 }], () => now),
    ).not.toThrow();
  });

  it('falls through when the bearer payload is not valid base64 JSON', () => {
    const now = Date.UTC(2026, 0, 1);
    const got = deriveTokenExpiry('a.!!!.c', [], () => now);
    expect(Date.parse(got)).toBe(now + 7 * 24 * 60 * 60 * 1000);
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

  it('returns an empty string when nothing matches', () => {
    expect(collectCookieHeader([all[2]], 'x.sharepoint.com')).toBe('');
  });
});
