import { describe, expect, it } from 'vitest';

import { classifyStatus, isStaleDigest, SharepointHttpError } from '../src/http/errors';
import { describeFetchError, isRetryableNetworkError } from '../src/http/client';

const STALE =
  '{"error":{"message":"The security validation for this page is invalid and might be corrupted."}}';

describe('isStaleDigest', () => {
  it('detects the digest marker on a 403', () => {
    expect(isStaleDigest(403, STALE)).toBe(true);
  });

  it('is case-insensitive', () => {
    expect(isStaleDigest(403, 'SECURITY VALIDATION FOR THIS PAGE IS INVALID')).toBe(true);
  });

  it('does not fire on a plain 403 access denial', () => {
    expect(isStaleDigest(403, '{"error":{"code":"accessDenied"}}')).toBe(false);
  });

  it('does not fire on a 401 even with the marker present', () => {
    expect(isStaleDigest(401, STALE)).toBe(false);
  });

  it('does not fire on an empty body', () => {
    expect(isStaleDigest(403, '')).toBe(false);
  });
});

describe('classifyStatus', () => {
  it.each([
    [401, '', 'AUTH_REQUIRED'],
    [403, '{"error":"accessDenied"}', 'ACCESS_DENIED'],
    [404, '', 'NOT_FOUND'],
    [410, '', 'NOT_FOUND'],
    [423, '', 'LOCKED'],
    [507, '', 'QUOTA_EXCEEDED'],
    [500, '', 'UPSTREAM'],
    [418, '', 'UPSTREAM'],
  ])('maps %i to %s', (status, body, expected) => {
    expect(classifyStatus(status as number, body as string)).toBe(expected);
  });

  it('classifies a stale-digest 403 as UPSTREAM, not ACCESS_DENIED', () => {
    // The retry path handles it; reporting access_denied would send the
    // operator after a permissions problem that does not exist.
    expect(classifyStatus(403, STALE)).toBe('UPSTREAM');
  });
});

describe('SharepointHttpError', () => {
  it('carries status, url and a derived code', () => {
    const e = new SharepointHttpError(404, 'https://x.sharepoint.com/a', 'gone');
    expect(e.status).toBe(404);
    expect(e.code).toBe('NOT_FOUND');
    expect(e.name).toBe('SharepointHttpError');
  });

  it('keeps the response body out of the message', () => {
    const e = new SharepointHttpError(500, 'https://x.sharepoint.com/a', 'SENSITIVE-BODY');
    expect(e.message).not.toContain('SENSITIVE-BODY');
    expect(e.body).toBe('SENSITIVE-BODY');
  });

  it('is instanceof Error for catch-site checks', () => {
    expect(new SharepointHttpError(500, 'u', '')).toBeInstanceOf(Error);
  });
});

describe('network error diagnosis', () => {
  // Node's fetch throws TypeError("fetch failed") and buries the reason on
  // .cause. That cost real debugging time when a SharePoint front-end started
  // taking longer to connect than undici's fixed 10s default allowed.
  const connectTimeout = Object.assign(new TypeError('fetch failed'), {
    cause: Object.assign(new Error('Connect Timeout Error'), {
      code: 'UND_ERR_CONNECT_TIMEOUT',
    }),
  });

  it('unwraps the cause instead of reporting a bare "fetch failed"', () => {
    const msg = describeFetchError(connectTimeout);
    expect(msg).toContain('Connect Timeout Error');
    expect(msg).toContain('UND_ERR_CONNECT_TIMEOUT');
  });

  it('handles an error with no cause', () => {
    expect(describeFetchError(new Error('plain'))).toBe('plain');
  });

  it('treats connection-level failures as retryable', () => {
    expect(isRetryableNetworkError(connectTimeout)).toBe(true);
    expect(isRetryableNetworkError({ code: 'ECONNRESET' })).toBe(true);
  });

  it('does not treat an application error as retryable', () => {
    // Once the server has answered, the status is the answer: retrying a 403
    // would just repeat it.
    expect(isRetryableNetworkError(new Error('boom'))).toBe(false);
    expect(isRetryableNetworkError({ code: 'ERR_SOMETHING_ELSE' })).toBe(false);
  });
});
