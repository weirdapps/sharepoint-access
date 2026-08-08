// src/commands/get.ts

import * as fs from 'node:fs';
import * as path from 'node:path';

import type { SharepointClient } from '../http/client';
import { CliError } from '../config/errors';
import { fileApi, normalizeServerRelative, splitParentLeaf } from '../sharepoint/paths';

export interface GetResult {
  source: string;
  size: number;
  contentType: string;
  /** Set when the bytes were written to disk rather than returned. */
  outPath?: string;
  /** Filename from Content-Disposition, when the server supplied one. */
  filename?: string;
}

type Reader = Pick<SharepointClient, 'getBinary'>;

function isAbsoluteUrl(s: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(s);
}

export async function runGet(
  client: Reader,
  pathOrUrl: string,
  outPath?: string,
): Promise<GetResult> {
  // An absolute URL goes through unchanged: the client host-checks it before
  // attaching cookies. A server-relative path gets the $value accessor.
  const target = isAbsoluteUrl(pathOrUrl)
    ? pathOrUrl
    : `${fileApi(normalizeServerRelative(pathOrUrl))}/$value`;

  const res = await client.getBinary(target);

  if (outPath) {
    const dir = path.dirname(path.resolve(outPath));
    try {
      await fs.promises.mkdir(dir, { recursive: true });
      await fs.promises.writeFile(path.resolve(outPath), res.bytes);
    } catch (err) {
      throw new CliError('IO', `failed to write "${outPath}": ${(err as Error).message}`);
    }
  }

  let filename = res.filename;
  if (!filename && !isAbsoluteUrl(pathOrUrl)) {
    filename = splitParentLeaf(pathOrUrl).leaf;
  }

  return {
    source: pathOrUrl,
    size: res.bytes.length,
    contentType: res.contentType,
    ...(outPath ? { outPath: path.resolve(outPath) } : {}),
    ...(filename ? { filename } : {}),
  };
}
