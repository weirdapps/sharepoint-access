// src/config/errors.ts
//
// Typed CLI errors and their mapping to the shared exit-code contract.

import { ExitCode, ExitWithCode, type ExitCodeValue } from '../util/exit-codes';

export type ErrorCode =
  | 'CONFIG_MISSING'
  | 'CONFIG_INVALID'
  | 'AUTH_REQUIRED'
  | 'ACCESS_DENIED'
  | 'NOT_FOUND'
  | 'LOCKED'
  | 'QUOTA_EXCEEDED'
  | 'UPSTREAM'
  | 'TIMEOUT'
  | 'IO';

const EXIT_BY_CODE: Record<ErrorCode, ExitCodeValue> = {
  CONFIG_MISSING: ExitCode.Config,
  CONFIG_INVALID: ExitCode.Config,
  AUTH_REQUIRED: ExitCode.AuthRequired,
  ACCESS_DENIED: ExitCode.Upstream,
  NOT_FOUND: ExitCode.Upstream,
  LOCKED: ExitCode.Upstream,
  QUOTA_EXCEEDED: ExitCode.Upstream,
  UPSTREAM: ExitCode.Upstream,
  TIMEOUT: ExitCode.Upstream,
  IO: ExitCode.Io,
};

export class CliError extends Error {
  constructor(
    public readonly code: ErrorCode,
    message: string,
    public readonly detail?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'CliError';
    Object.setPrototypeOf(this, new.target.prototype);
  }

  toExit(): ExitWithCode {
    return new ExitWithCode(EXIT_BY_CODE[this.code], {
      code: this.code.toLowerCase(),
      message: this.message,
      ...(this.detail ?? {}),
    });
  }
}

export function exitCodeFor(code: ErrorCode): ExitCodeValue {
  return EXIT_BY_CODE[code];
}
