// src/commands/mkdir.ts

import type { SharepointClient } from '../http/client';
import { foldersAddApi, normalizeServerRelative, PathError } from '../sharepoint/paths';

export interface MkdirResult {
  serverRelativeUrl: string;
  created: true;
}

type Writer = Pick<SharepointClient, 'postJson'>;

/**
 * Creates one folder. Parents must already exist: there is deliberately no
 * implicit -p, so a typo cannot scatter junk folders through a live library.
 */
export async function runMkdir(client: Writer, path: string, site?: string): Promise<MkdirResult> {
  const norm = normalizeServerRelative(path);
  if (norm === '/') throw new PathError('cannot create the root folder');
  await client.postJson(foldersAddApi(norm, site));
  return { serverRelativeUrl: norm, created: true };
}
