// src/util/exit-codes.ts
//
// Mirrors teams-access/src/util/exit-codes.ts exactly for codes 0-6 so callers
// and cron wrappers can treat all three CLIs identically: 4 always means
// "credentials are gone, re-authenticate", 5 always means "the far end
// misbehaved". 7 is new here and exists only while the CLI is a scaffold.

export const ExitCode = {
  Success: 0,
  Internal: 1,
  InvalidInput: 2,
  Config: 3,
  AuthRequired: 4,
  Upstream: 5,
  Io: 6,
  /** Command exists in the design but is not built yet. Remove once phase 3 lands. */
  NotImplemented: 7,
} as const;
export type ExitCodeValue = (typeof ExitCode)[keyof typeof ExitCode];

export interface ErrorPayload {
  code: string;
  message?: string;
  [key: string]: unknown;
}

export class ExitWithCode extends Error {
  constructor(
    public readonly code: ExitCodeValue,
    public readonly payload: ErrorPayload,
  ) {
    super(payload.message ?? payload.code);
    this.name = 'ExitWithCode';
  }
}
