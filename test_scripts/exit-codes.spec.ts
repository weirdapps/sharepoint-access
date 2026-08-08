// test_scripts/exit-codes.spec.ts
//
// Locks the exit-code contract shared with outlook-cli and teams-cli. Cron
// wrappers branch on these numbers, so a rename or renumber here silently
// breaks callers in other repos. The literals are duplicated on purpose: if
// this test imported its expectations from the same source it verifies, it
// would pass no matter what the values became.

import { describe, expect, it } from 'vitest';

import { ExitCode, ExitWithCode } from '../src/util/exit-codes';

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

  it('reserves 7 for the scaffold-only NotImplemented case', () => {
    // Remove this expectation, and the enum member, once phase 3 lands.
    // Tracked as P4 in "Issues - Pending Items.md".
    expect(ExitCode.NotImplemented).toBe(7);
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
    const err = new ExitWithCode(ExitCode.Upstream, { code: 'upstream_error' });
    expect(err.message).toBe('upstream_error');
  });
});
