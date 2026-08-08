// test_scripts/exit-codes.spec.ts
//
// Locks the exit-code contract shared with outlook-cli and teams-cli. Cron
// wrappers branch on these numbers, so a rename or renumber here silently
// breaks callers in other repos. The literals are duplicated on purpose: if
// this test imported its expectations from the same source it verifies, it
// would pass no matter what the values became.

import { describe, expect, it } from 'vitest';

import { ExitCode, ExitWithCode } from '../src/util/exit-codes';
import { CliError, exitCodeFor } from '../src/config/errors';

describe('ExitCode', () => {
  it('matches the values teams-access and outlook-access use for 0-6', () => {
    expect(ExitCode.Success).toBe(0);
    expect(ExitCode.Internal).toBe(1);
    expect(ExitCode.InvalidInput).toBe(2);
    expect(ExitCode.Config).toBe(3);
    expect(ExitCode.AuthRequired).toBe(4);
    expect(ExitCode.Upstream).toBe(5);
    expect(ExitCode.Io).toBe(6);
  });

  it('no longer carries the scaffold-only NotImplemented code', () => {
    expect((ExitCode as Record<string, number>).NotImplemented).toBeUndefined();
  });
});

describe('ExitWithCode', () => {
  it('takes its message from the payload', () => {
    const err = new ExitWithCode(ExitCode.AuthRequired, {
      code: 'auth_required',
      message: 'session expired',
    });
    expect(err.message).toBe('session expired');
    expect(err.code).toBe(4);
    expect(err.name).toBe('ExitWithCode');
  });

  it('falls back to the payload code when no message is given', () => {
    expect(new ExitWithCode(ExitCode.Upstream, { code: 'upstream_error' }).message).toBe(
      'upstream_error',
    );
  });
});

describe('exitCodeFor', () => {
  it('maps auth failures to 4 so cron can branch on re-authentication', () => {
    expect(exitCodeFor('AUTH_REQUIRED')).toBe(4);
  });

  it('maps every upstream-ish failure to 5', () => {
    for (const c of [
      'ACCESS_DENIED',
      'NOT_FOUND',
      'LOCKED',
      'QUOTA_EXCEEDED',
      'UPSTREAM',
      'TIMEOUT',
    ] as const) {
      expect(exitCodeFor(c)).toBe(5);
    }
  });

  it('maps config problems to 3 and IO to 6', () => {
    expect(exitCodeFor('CONFIG_MISSING')).toBe(3);
    expect(exitCodeFor('CONFIG_INVALID')).toBe(3);
    expect(exitCodeFor('IO')).toBe(6);
  });
});

describe('CliError', () => {
  it('converts to an ExitWithCode carrying a lowercased code', () => {
    const e = new CliError('AUTH_REQUIRED', 'gone').toExit();
    expect(e.code).toBe(4);
    expect(e.payload.code).toBe('auth_required');
  });

  it('merges detail into the payload', () => {
    const e = new CliError('NOT_FOUND', 'missing', { path: '/a' }).toExit();
    expect(e.payload.path).toBe('/a');
  });
});
