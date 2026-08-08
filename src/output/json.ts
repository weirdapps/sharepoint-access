// src/output/json.ts
//
// Every command prints JSON to stdout and JSON errors to stderr, matching
// outlook-cli and teams-cli so the same wrappers can parse all three.

import { CliError, exitCodeFor } from '../config/errors';
import { SharepointHttpError } from '../http/errors';
import { ExitCode, ExitWithCode, type ExitCodeValue } from '../util/exit-codes';

export function emit(value: unknown): void {
  process.stdout.write(JSON.stringify(value, null, 2) + '\n');
}

export function emitError(payload: Record<string, unknown>): void {
  process.stderr.write(JSON.stringify(payload) + '\n');
}

/**
 * Map any thrown value to an exit code and a JSON error payload.
 *
 * Deliberately does NOT include the HTTP response body: it can echo request
 * content, and this text is what lands in cron logs.
 */
export function toExit(err: unknown): { code: ExitCodeValue; payload: Record<string, unknown> } {
  if (err instanceof CliError) {
    return {
      code: exitCodeFor(err.code),
      payload: { error: err.code.toLowerCase(), message: err.message, ...(err.detail ?? {}) },
    };
  }
  if (err instanceof SharepointHttpError) {
    return {
      code: exitCodeFor(err.code),
      payload: {
        error: err.code.toLowerCase(),
        message: err.message,
        status: err.status,
        ...(err.code === 'AUTH_REQUIRED'
          ? { hint: 'run "sharepoint-cli auth-renew", or "login" if that fails' }
          : {}),
      },
    };
  }
  if (err instanceof ExitWithCode) {
    return { code: err.code, payload: { ...err.payload } };
  }
  return {
    code: ExitCode.Internal,
    payload: { error: 'internal', message: (err as Error)?.message ?? String(err) },
  };
}
