import { describe, expect, it } from 'vitest';

import { deriveTokenExpiry, collectCookieHeader } from '../src/auth/capture';

describe('deriveTokenExpiry', () => {
  it('reports the FedAuth cookie expiry', () => {
    const secs = Math.floor(Date.UTC(2029, 0, 1) / 1000);
    expect(deriveTokenExpiry([{ name: 'FedAuth', expires: secs }])).toBe(
      new Date(secs * 1000).toISOString(),
    );
  });

  it('prefers FedAuth over rtFa', () => {
    const a = Math.floor(Date.UTC(2029, 0, 1) / 1000);
    const b = Math.floor(Date.UTC(2028, 0, 1) / 1000);
    expect(
      deriveTokenExpiry([
        { name: 'rtFa', expires: b },
        { name: 'FedAuth', expires: a },
      ]),
    ).toBe(new Date(a * 1000).toISOString());
  });

  it('matches the cookie name case-insensitively', () => {
    const secs = Math.floor(Date.UTC(2029, 0, 1) / 1000);
    expect(deriveTokenExpiry([{ name: 'fedauth', expires: secs }])).toBe(
      new Date(secs * 1000).toISOString(),
    );
  });

  it('uses the conservative window for session cookies', () => {
    const now = Date.UTC(2026, 0, 1);
    const got = Date.parse(deriveTokenExpiry([{ name: 'FedAuth', expires: -1 }], () => now));
    expect(got).toBe(now + 7 * 24 * 60 * 60 * 1000);
  });

  it('uses the conservative window when no cookie carries an expiry', () => {
    const now = Date.UTC(2026, 0, 1);
    expect(Date.parse(deriveTokenExpiry([], () => now))).toBe(now + 7 * 24 * 60 * 60 * 1000);
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
