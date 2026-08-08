// src/session/store.ts
//
// Atomic, permission-tight IO for the session file. The file holds live
// credentials, so it is written 0600 inside a 0700 directory, and via a
// temp-then-rename so a crash mid-write cannot leave a truncated session.

import * as fs from 'node:fs';
import * as path from 'node:path';

import { parseSession, serializeSession, type SharepointSession } from './schema';

export { parseSession, serializeSession, SessionParseError } from './schema';
export type { SharepointSession } from './schema';

export async function loadSession(filePath: string): Promise<SharepointSession | null> {
  let data: string;
  try {
    data = await fs.promises.readFile(filePath, 'utf8');
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return null;
    throw err;
  }
  // A corrupt file is NOT the same as an absent one: returning null would
  // silently trigger a re-login and mask on-disk damage.
  return parseSession(data);
}

export async function saveSession(filePath: string, session: SharepointSession): Promise<void> {
  const dir = path.dirname(filePath);
  await fs.promises.mkdir(dir, { recursive: true, mode: 0o700 });
  try {
    await fs.promises.chmod(dir, 0o700);
  } catch {
    // Tolerated on pre-existing directories we do not own the inode for.
  }

  const tmp = `${filePath}.tmp`;
  await fs.promises.writeFile(tmp, serializeSession(session), { mode: 0o600 });
  // writeFile honours `mode` only when it creates the file; an existing tmp
  // keeps its old mode, so set it explicitly.
  await fs.promises.chmod(tmp, 0o600);
  await fs.promises.rename(tmp, filePath);
}
