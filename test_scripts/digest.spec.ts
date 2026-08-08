import { describe, expect, it, vi } from 'vitest';

import { DigestCache } from '../src/sharepoint/digest';

function clientReturning(
  ...values: Array<{ FormDigestValue?: string; FormDigestTimeoutSeconds?: number }>
) {
  const contextInfo = vi.fn();
  for (const v of values) contextInfo.mockResolvedValueOnce(v);
  return { contextInfo };
}

describe('DigestCache', () => {
  it('fetches once and caches', async () => {
    const c = clientReturning({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 });
    const d = new DigestCache(c);
    expect(await d.get()).toBe('D1');
    expect(await d.get()).toBe('D1');
    expect(c.contextInfo).toHaveBeenCalledTimes(1);
  });

  it('uses the dedicated contextInfo call, never the digest-carrying postJson', async () => {
    // Routing this through postJson would recurse: postJson asks the digest
    // provider for a header, and the provider's only source is this call.
    const c = clientReturning({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 });
    await new DigestCache(c).get();
    expect(c.contextInfo).toHaveBeenCalledTimes(1);
  });

  it('refetches after expiry, applying the 60s safety margin', async () => {
    const c = clientReturning(
      { FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 },
      { FormDigestValue: 'D2', FormDigestTimeoutSeconds: 1800 },
    );
    let t = 1_000_000;
    const d = new DigestCache(c, () => t);
    expect(await d.get()).toBe('D1');
    // 1740s = 1800 - 60. Just inside the boundary the cache must still hold.
    t += 1_739_000;
    expect(await d.get()).toBe('D1');
    t += 2_000;
    expect(await d.get()).toBe('D2');
  });

  it('refetches when forced, even inside the validity window', async () => {
    const c = clientReturning(
      { FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 },
      { FormDigestValue: 'D2', FormDigestTimeoutSeconds: 1800 },
    );
    const d = new DigestCache(c, () => 1_000_000);
    expect(await d.get()).toBe('D1');
    expect(await d.get(true)).toBe('D2');
  });

  it('treats a missing timeout as a conservative 60s of validity', async () => {
    const c = clientReturning(
      { FormDigestValue: 'D1' },
      { FormDigestValue: 'D2', FormDigestTimeoutSeconds: 1800 },
    );
    let t = 1_000_000;
    const d = new DigestCache(c, () => t);
    expect(await d.get()).toBe('D1');
    t += 1_000;
    expect(await d.get()).toBe('D2');
  });

  it('raises when contextinfo returns no digest value', async () => {
    const c = clientReturning({ FormDigestTimeoutSeconds: 1800 });
    await expect(new DigestCache(c).get()).rejects.toThrowError(/FormDigestValue/);
  });

  it('raises when contextinfo returns an empty digest value', async () => {
    const c = clientReturning({ FormDigestValue: '', FormDigestTimeoutSeconds: 1800 });
    await expect(new DigestCache(c).get()).rejects.toThrowError(/FormDigestValue/);
  });

  it('does not issue concurrent fetches for simultaneous callers', async () => {
    const contextInfo = vi
      .fn()
      .mockImplementation(
        () =>
          new Promise((r) =>
            setTimeout(() => r({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 }), 10),
          ),
      );
    const d = new DigestCache({ contextInfo });
    const [a, b] = await Promise.all([d.get(), d.get()]);
    expect(a).toBe('D1');
    expect(b).toBe('D1');
    expect(contextInfo).toHaveBeenCalledTimes(1);
  });

  it('recovers after a failed fetch rather than caching the rejection', async () => {
    const contextInfo = vi
      .fn()
      .mockRejectedValueOnce(new Error('boom'))
      .mockResolvedValueOnce({ FormDigestValue: 'D1', FormDigestTimeoutSeconds: 1800 });
    const d = new DigestCache({ contextInfo });
    await expect(d.get()).rejects.toThrowError('boom');
    expect(await d.get()).toBe('D1');
  });
});
