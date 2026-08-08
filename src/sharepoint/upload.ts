// src/sharepoint/upload.ts
//
// Size-routed upload. Below the threshold a single addUsingPath request;
// at or above it a chunked StartUpload/ContinueUpload/FinishUpload session.

import { randomUUID } from 'node:crypto';

import type { SharepointClient } from '../http/client';
import { encodeLeaf, fileApi, folderApi, joinPath, normalizeServerRelative } from './paths';

/** At or above this size, use a chunked session. */
export const CHUNK_THRESHOLD_BYTES = 10 * 1024 * 1024;
export const CHUNK_SIZE_BYTES = 10 * 1024 * 1024;

export interface UploadResult {
  serverRelativeUrl: string;
  size: number;
  chunked: boolean;
}

type Writer = Pick<SharepointClient, 'postBinary' | 'postJson'>;

function addUsingPath(folder: string, leaf: string, overwrite: boolean): string {
  return `${folderApi(folder)}/Files/addUsingPath(DecodedUrl='${encodeLeaf(leaf)}',overwrite=${overwrite})`;
}

export async function uploadFile(
  client: Writer,
  bytes: Buffer,
  remoteFolder: string,
  leaf: string,
  overwrite: boolean,
): Promise<UploadResult> {
  const folder = normalizeServerRelative(remoteFolder);
  const target = joinPath(folder, leaf);

  if (bytes.length < CHUNK_THRESHOLD_BYTES) {
    await client.postBinary(addUsingPath(folder, leaf, overwrite), bytes);
    return { serverRelativeUrl: target, size: bytes.length, chunked: false };
  }

  // SPO requires the target to exist before a chunked session can start.
  await client.postBinary(addUsingPath(folder, leaf, overwrite), Buffer.alloc(0));

  const uploadId = randomUUID();
  const file = fileApi(target);
  try {
    const first = bytes.subarray(0, CHUNK_SIZE_BYTES);
    await client.postBinary(`${file}/StartUpload(uploadId=guid'${uploadId}')`, Buffer.from(first));
    let offset = first.length;

    // The final chunk always goes to FinishUpload, so stop short of it here
    // even when the size divides evenly: a zero-byte ContinueUpload would
    // otherwise be emitted for an exact multiple.
    while (bytes.length - offset > CHUNK_SIZE_BYTES) {
      const chunk = bytes.subarray(offset, offset + CHUNK_SIZE_BYTES);
      await client.postBinary(
        `${file}/ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`,
        Buffer.from(chunk),
      );
      offset += chunk.length;
    }

    await client.postBinary(
      `${file}/FinishUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`,
      Buffer.from(bytes.subarray(offset)),
    );
    return { serverRelativeUrl: target, size: bytes.length, chunked: true };
  } catch (err) {
    // Leave no partial file behind: a retry must start clean rather than
    // resume into an inconsistent file.
    try {
      await client.postJson(`${file}/recycle()`);
    } catch {
      // Best-effort. Must never mask the original failure.
    }
    throw err;
  }
}
