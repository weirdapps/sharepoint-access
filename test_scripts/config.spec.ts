import { describe, expect, it } from 'vitest';

import { loadConfig } from '../src/config/load';
import { CliError } from '../src/config/errors';

describe('loadConfig', () => {
  it('raises CONFIG_MISSING when host is absent everywhere', () => {
    expect(() => loadConfig({}, {})).toThrowError(CliError);
    try {
      loadConfig({}, {});
    } catch (e) {
      expect((e as CliError).code).toBe('CONFIG_MISSING');
    }
  });

  it('takes host from the flag ahead of the env var', () => {
    const c = loadConfig({ host: 'a.sharepoint.com' }, { SHAREPOINT_CLI_HOST: 'b.sharepoint.com' });
    expect(c.host).toBe('a.sharepoint.com');
  });

  it('falls back to the env var for host', () => {
    const c = loadConfig({}, { SHAREPOINT_CLI_HOST: 'b.sharepoint.com' });
    expect(c.host).toBe('b.sharepoint.com');
  });

  it('rejects a host that is not a sharepoint.com domain', () => {
    expect(() => loadConfig({ host: 'evil.example.com' }, {})).toThrowError(/sharepoint\.com/);
  });

  it('rejects a host with sharepoint.com as a prefix of a longer domain', () => {
    expect(() => loadConfig({ host: 'x.sharepoint.com.evil.com' }, {})).toThrowError(
      /sharepoint\.com/,
    );
  });

  it('rejects a host carrying a path', () => {
    expect(() => loadConfig({ host: 'x.sharepoint.com/../evil' }, {})).toThrowError(
      /sharepoint\.com/,
    );
  });

  it('rejects a host carrying credentials or a port', () => {
    expect(() => loadConfig({ host: 'evil.com@x.sharepoint.com' }, {})).toThrowError(/sharepoint/);
    expect(() => loadConfig({ host: 'x.sharepoint.com:8443' }, {})).toThrowError(/sharepoint/);
  });

  it('accepts the -my OneDrive for Business host', () => {
    const c = loadConfig({ host: 'x-my.sharepoint.com' }, {});
    expect(c.host).toBe('x-my.sharepoint.com');
  });

  it('lowercases the host', () => {
    expect(loadConfig({ host: 'X.SharePoint.CoM' }, {}).host).toBe('x.sharepoint.com');
  });

  it('applies documented defaults for the four plumbing settings', () => {
    const c = loadConfig({ host: 'x.sharepoint.com' }, {});
    expect(c.httpTimeoutMs).toBe(30000);
    expect(c.loginTimeoutMs).toBe(300000);
    expect(c.renewTimeoutMs).toBe(30000);
    expect(c.chromeChannel).toBe('chrome');
  });

  it('lets env vars override plumbing defaults', () => {
    const c = loadConfig({ host: 'x.sharepoint.com' }, { SHAREPOINT_CLI_HTTP_TIMEOUT_MS: '1234' });
    expect(c.httpTimeoutMs).toBe(1234);
  });

  it('rejects a non-numeric timeout rather than silently defaulting', () => {
    expect(() =>
      loadConfig({ host: 'x.sharepoint.com' }, { SHAREPOINT_CLI_HTTP_TIMEOUT_MS: 'abc' }),
    ).toThrowError(/SHAREPOINT_CLI_HTTP_TIMEOUT_MS/);
  });

  it('rejects a zero or negative timeout', () => {
    expect(() =>
      loadConfig({ host: 'x.sharepoint.com' }, { SHAREPOINT_CLI_HTTP_TIMEOUT_MS: '0' }),
    ).toThrowError(/positive/);
  });

  it('derives state paths under a single directory', () => {
    const c = loadConfig({ host: 'x.sharepoint.com' }, {});
    expect(c.sessionPath).toMatch(/\.sharepoint-cli\/session\.json$/);
    expect(c.profileDir).toMatch(/\.sharepoint-cli\/playwright-profile$/);
    expect(c.lockPath).toMatch(/\.sharepoint-cli\/\.browser\.lock$/);
  });
});
