// src/commands/put.ts

import * as fs from 'node:fs';
import * as path from 'node:path';

import type { SharepointClient } from '../http/client';
import { CliError } from '../config/errors';
import { uploadFile, type UploadResult } from '../sharepoint/upload';

type Writer = Pick<SharepointClient, 'postBinary' | 'postJson'>;

export async function runPut(
  client: Writer,
  localPath: string,
  remoteFolder: string,
  overwrite: boolean,
): Promise<UploadResult> {
  let bytes: Buffer;
  try {
    bytes = await fs.promises.readFile(localPath);
  } catch (err) {
    const e = err as NodeJS.ErrnoException;
    throw new CliError(
      'IO',
      e.code === 'ENOENT'
        ? `local file not found: "${localPath}"`
        : `failed to read "${localPath}": ${e.message}`,
    );
  }
  const leaf = path.basename(localPath);
  if (leaf.length === 0 || leaf === '.' || leaf === '..') {
    throw new CliError('CONFIG_INVALID', `cannot derive a filename from "${localPath}"`);
  }
  return uploadFile(client, bytes, remoteFolder, leaf, overwrite);
}
