// src/commands/get.ts

import * as fs from 'node:fs';
import * as path from 'node:path';

import type { SharepointClient } from '../http/client';
import { CliError } from '../config/errors';
import { classifyLink, filePathOf, parseViewerPage } from '../sharepoint/links';
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
type Binary = Awaited<ReturnType<Reader['getBinary']>>;

function isAbsoluteUrl(s: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(s);
}

function isHtml(res: Binary): boolean {
  if (/text\/html/i.test(res.contentType)) return true;
  const head = res.bytes.subarray(0, 512).toString('utf8').trimStart().toLowerCase();
  return head.startsWith('<!doctype html') || head.startsWith('<html');
}

/**
 * The file behind an absolute URL. It is fetched as is: a document URL answers with the file.
 * A sharing or viewer link answers with the browser page, whose context names the file and the
 * web that owns it, so the file is fetched by ID from that web. A page that names no file is
 * NOT_A_FILE, never saved as the file.
 */
async function fetchLink(client: Reader, url: string): Promise<{ res: Binary; name?: string }> {
  const kind = classifyLink(url);
  if (kind === 'page') {
    throw new CliError('NOT_A_FILE', `not a file: ${url} is an intranet page; read it with "page"`);
  }
  const document = filePathOf(url);
  const leaf = document ? splitParentLeaf(document).leaf : undefined;
  const first = await client.getBinary(url);
  if (!isHtml(first)) return { res: first, ...(leaf ? { name: leaf } : {}) };
  // Only a file link's page is a viewer. Other pages (a OneDrive view) name the files they list.
  const context = kind === 'file' ? parseViewerPage(first.bytes.toString('utf8')) : null;
  if (!context) {
    throw new CliError('NOT_A_FILE', `not a file: ${url} returned a web page that names no file`);
  }
  const res = await client.getBinary(
    `${context.webAbsoluteUrl}/_api/web/GetFileById('${context.fileId}')/$value`,
  );
  const name = context.fileName ?? leaf;
  return { res, ...(name ? { name } : {}) };
}

export async function runGet(
  client: Reader,
  pathOrUrl: string,
  outPath?: string,
  site?: string,
): Promise<GetResult> {
  // An absolute URL is resolved to its file (fetchLink); the client host-checks every URL
  // before attaching cookies. A server-relative path gets the $value accessor.
  const { res, name } = isAbsoluteUrl(pathOrUrl)
    ? await fetchLink(client, pathOrUrl)
    : {
        res: await client.getBinary(`${fileApi(normalizeServerRelative(pathOrUrl), site)}/$value`),
      };

  if (outPath) {
    const dir = path.dirname(path.resolve(outPath));
    try {
      await fs.promises.mkdir(dir, { recursive: true });
      await fs.promises.writeFile(path.resolve(outPath), res.bytes);
    } catch (err) {
      throw new CliError('IO', `failed to write "${outPath}": ${(err as Error).message}`);
    }
  }

  let filename = res.filename ?? name;
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
