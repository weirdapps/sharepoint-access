// src/http/errors.ts

import type { ErrorCode } from '../config/errors';

/**
 * SharePoint answers a STALE DIGEST with 403, not 401. Without this
 * discriminator every expired write looks like a permissions failure and
 * sends the operator after a problem that does not exist.
 */
const STALE_DIGEST_MARKER = 'security validation for this page is invalid';

export function isStaleDigest(status: number, body: string): boolean {
  return status === 403 && body.toLowerCase().includes(STALE_DIGEST_MARKER);
}

export function classifyStatus(status: number, body: string): ErrorCode {
  if (status === 401) return 'AUTH_REQUIRED';
  // A stale digest is recoverable plumbing, not a permissions problem: the
  // retry-once envelope handles it, so it must not be reported as denial.
  if (status === 403) return isStaleDigest(status, body) ? 'UPSTREAM' : 'ACCESS_DENIED';
  if (status === 404 || status === 410) return 'NOT_FOUND';
  if (status === 423) return 'LOCKED';
  if (status === 507) return 'QUOTA_EXCEEDED';
  return 'UPSTREAM';
}

export class SharepointHttpError extends Error {
  public readonly code: ErrorCode;

  constructor(
    public readonly status: number,
    public readonly url: string,
    public readonly body: string,
  ) {
    // Message carries status and URL only. The body may echo request content
    // and callers log messages, so it stays on the property, not in the text.
    super(`SharePoint ${status} for ${url}`);
    this.name = 'SharepointHttpError';
    this.code = classifyStatus(status, body);
    Object.setPrototypeOf(this, new.target.prototype);
  }
}
