import { describe, expect, it, vi } from 'vitest';

import { runAuthCheck } from '../src/commands/auth-check';
import { SharepointHttpError } from '../src/http/errors';

// SharepointHttpError deliberately keeps the response body OFF its message,
// because callers log messages and a body can echo request content. That rule
// is right in general and starved the operator here: on 2026-09-10 all three
// probes reported the bare string "SharePoint 500" and nothing else, so the
// estate page told its reader SharePoint was broken while the real fault was a
// credential of ours. The body SharePoint actually sent that day was "The token
// being parsed does not have an issuer."
//
// auth-check is the one caller where including it is unconditionally safe: its
// three URLs are fixed, carry no user content, and cannot echo anything back.
describe('probe detail carries the server explanation', () => {
  const failing = (err: Error) => ({
    getJson: vi.fn().mockRejectedValue(err),
    contextInfo: vi.fn().mockRejectedValue(err),
  });

  it('appends the response body to the probe detail', async () => {
    const err = new SharepointHttpError(
      500,
      'https://x.sharepoint.com/_api/web',
      'The token being parsed does not have an issuer.',
    );
    const r = await runAuthCheck(failing(err) as never);
    const read = r.probes.find((p) => p.name === 'read');
    expect(read?.detail).toContain('SharePoint 500');
    expect(read?.detail).toContain('does not have an issuer');
  });

  it('collapses a multi-line body to one line and bounds its length', async () => {
    const err = new SharepointHttpError(
      500,
      'https://x.sharepoint.com/_api/web',
      'a\nb'.padEnd(900, 'x'),
    );
    const r = await runAuthCheck(failing(err) as never);
    const detail = r.probes.find((p) => p.name === 'read')?.detail ?? '';
    expect(detail).not.toContain('\n');
    expect(detail.length).toBeLessThan(400);
  });

  it('leaves the detail alone when the server sent no body', async () => {
    const err = new SharepointHttpError(500, 'https://x.sharepoint.com/_api/web', '');
    const r = await runAuthCheck(failing(err) as never);
    expect(r.probes.find((p) => p.name === 'read')?.detail).toBe(err.message);
  });

  it('leaves a non-HTTP error untouched', async () => {
    const r = await runAuthCheck(failing(new Error('request timed out')) as never);
    expect(r.probes.find((p) => p.name === 'read')?.detail).toBe('request timed out');
  });
});

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
