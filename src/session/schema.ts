// src/session/schema.ts
//
// Shape of ~/.sharepoint-cli/session.json. Ported from
// outlook-access/src/session/sharepoint-schema.ts. Pure parse/serialize only;
// the IO lives in store.ts.

export interface SharepointSession {
  version: 1;
  /** SharePoint host, e.g. "<tenant>.sharepoint.com" or "<tenant>-my.sharepoint.com". */
  host: string;
  /**
   * Bearer token (no "Bearer " prefix). Optional: cookie-authenticated tenants
   * never emit one, and the FedAuth/rtFa cookies authorise on their own. Sent
   * as an Authorization header when present.
   */
  bearer?: string;
  /** Serialized cookie header value, e.g. "FedAuth=…; rtFa=…". */
  cookies: string;
  /** ISO-8601 UTC timestamp of capture. */
  capturedAt: string;
  /** ISO-8601 UTC. From the JWT exp when a bearer exists, else cookie expiry. */
  tokenExpiresAt: string;
}

export class SessionParseError extends Error {
  constructor(msg: string) {
    super(msg);
    this.name = 'SessionParseError';
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

export function parseSession(json: string): SharepointSession {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch (e) {
    throw new SessionParseError(`Invalid JSON: ${(e as Error).message}`);
  }
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    throw new SessionParseError('Expected JSON object');
  }
  const obj = raw as Record<string, unknown>;
  if (obj.version !== 1) {
    throw new SessionParseError(`Unsupported version: ${String(obj.version)}`);
  }
  for (const key of ['host', 'cookies', 'capturedAt', 'tokenExpiresAt']) {
    if (typeof obj[key] !== 'string' || (obj[key] as string).length === 0) {
      throw new SessionParseError(`Missing or invalid "${key}"`);
    }
  }
  if (obj.bearer !== undefined && typeof obj.bearer !== 'string') {
    throw new SessionParseError('Invalid "bearer" (must be a string when present)');
  }
  return obj as unknown as SharepointSession;
}

export function serializeSession(s: SharepointSession): string {
  return JSON.stringify(s, null, 2);
}
