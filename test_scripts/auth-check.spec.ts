import { describe, expect, it, vi } from 'vitest';

import { runAuthCheck } from '../src/commands/auth-check';

const okClient = () => ({
  getJson: vi.fn().mockResolvedValue({ Title: 'T' }),
  contextInfo: vi.fn().mockResolvedValue({ FormDigestValue: 'D', FormDigestTimeoutSeconds: 1800 }),
});

describe('runAuthCheck', () => {
  it('probes read, write and search, not just one', async () => {
    const c = okClient();
    const r = await runAuthCheck(c as never);
    expect(r.probes.map((p) => p.name).sort()).toEqual(['read', 'search', 'write']);
    expect(r.overall).toBe('ok');
  });

  it('probes contextinfo for the write surface, web-scoped', async () => {
    const c = okClient();
    await runAuthCheck(c as never);
    expect(c.contextInfo).toHaveBeenCalledWith('');
  });

  it('probes the requested sub-site web, not the host root', async () => {
    // A green root-web probe says nothing about whether a sub-site is writable:
    // digests and folder access are both web-scoped.
    const c = okClient();
    await runAuthCheck(c as never, '/personal/u');
    expect(c.contextInfo).toHaveBeenCalledWith('/personal/u');
    expect(
      c.getJson.mock.calls.every((call) => (call[0] as string).startsWith('/personal/u/')),
    ).toBe(true);
  });

  it('reports degraded when only the search probe fails', async () => {
    const c = okClient();
    c.getJson = vi
      .fn()
      .mockImplementation((u: string) =>
        u.includes('/_api/search')
          ? Promise.reject(new Error('no search'))
          : Promise.resolve({ Title: 'T' }),
      );
    const r = await runAuthCheck(c as never);
    expect(r.overall).toBe('degraded');
    expect(r.probes.find((p) => p.name === 'search')?.ok).toBe(false);
    expect(r.probes.find((p) => p.name === 'read')?.ok).toBe(true);
  });

  it('reports degraded when only the write probe fails, which the teams bug would have missed', async () => {
    const c = okClient();
    c.contextInfo = vi.fn().mockRejectedValue(new Error('no digest'));
    const r = await runAuthCheck(c as never);
    expect(r.overall).toBe('degraded');
    expect(r.probes.find((p) => p.name === 'write')?.ok).toBe(false);
  });

  it('reports broken when the read probe fails', async () => {
    const c = okClient();
    c.getJson = vi.fn().mockRejectedValue(new Error('401'));
    c.contextInfo = vi.fn().mockRejectedValue(new Error('401'));
    expect((await runAuthCheck(c as never)).overall).toBe('broken');
  });

  it('reports broken on a read failure even when write and search pass', async () => {
    const c = okClient();
    c.getJson = vi
      .fn()
      .mockImplementation((u: string) =>
        u.includes('/_api/search') ? Promise.resolve({}) : Promise.reject(new Error('read denied')),
      );
    expect((await runAuthCheck(c as never)).overall).toBe('broken');
  });

  it('records a duration for every probe', async () => {
    const r = await runAuthCheck(okClient() as never);
    for (const p of r.probes) expect(typeof p.durationMs).toBe('number');
  });

  it('carries the failure detail rather than swallowing it', async () => {
    const c = okClient();
    c.contextInfo = vi.fn().mockRejectedValue(new Error('digest exploded'));
    const r = await runAuthCheck(c as never);
    expect(r.probes.find((p) => p.name === 'write')?.detail).toBe('digest exploded');
  });

  it('never throws, so cron always gets a report', async () => {
    const c = {
      getJson: vi.fn().mockRejectedValue(new Error('x')),
      contextInfo: vi.fn().mockRejectedValue(new Error('y')),
    };
    await expect(runAuthCheck(c as never)).resolves.toBeDefined();
  });
});
