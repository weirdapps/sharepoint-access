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
   * Serialized cookie header value, e.g. "FedAuth=…; rtFa=…".
   *
   * This is the ONLY credential. There used to be an optional `bearer` beside
   * it, scavenged from the page's own traffic during capture; it was dropped
   * because the client never sent it and it expired inside the hour while the
   * cookies stayed valid for days. Sessions written before that change still
   * carry the key, and parseSession ignores it rather than rejecting them.
   */
  cookies: string;
  /** ISO-8601 UTC timestamp of capture. */
  capturedAt: string;
  /** ISO-8601 UTC. The FedAuth (then rtFa) cookie expiry. */
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
  return obj as unknown as SharepointSession;
}

export function serializeSession(s: SharepointSession): string {
  return JSON.stringify(s, null, 2);
}
