// src/commands/get.ts

import * as fs from 'node:fs';
import * as path from 'node:path';

import type { SharepointClient } from '../http/client';
import { CliError } from '../config/errors';
import {
  classifyLink,
  filePathOf,
  isSingleFileView,
  pageWebOf,
  parseViewerPage,
  sourcedocOf,
  viewedPathOf,
} from '../sharepoint/links';
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

/** The server-relative path of the web that owns `path`: the page's own web when its path
 * prefixes the file's, else undefined, and fileApi derives it from the path. A OneDrive view
 * is served from the reader's own OneDrive, which does not own the file it shows. */
function owningWeb(path: string, pageWeb: string | undefined, host: string): string | undefined {
  if (!pageWeb) return undefined;
  const web = new URL(pageWeb);
  const webPath = decodeURIComponent(web.pathname).replace(/\/+$/, '');
  const owns = path.toLowerCase().startsWith(`${webPath.toLowerCase()}/`);
  return web.host === host && owns ? webPath : undefined;
}

/**
 * The file behind an absolute URL. It is fetched as is: a document URL answers with the file.
 * A sharing or viewer link answers with a page, which names the file one of two ways. A
 * OneDrive or library view carries the file's path in its id parameter: fetched by that path.
 * The viewer of one file carries its FileId (the landing's sourcedoc when there is one): fetched
 * by ID from the page's web. Any other page is NOT_A_FILE, never saved as the file.
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
  const notAFile = new CliError(
    'NOT_A_FILE',
    `not a file: ${url} returned a web page that names no file`,
  );
  // Only a file link's page can name its file. Other pages (a OneDrive view) list many.
  if (kind !== 'file') throw notAFile;
  const landing = first.url ?? url;
  const page = first.bytes.toString('utf8');
  const viewed = viewedPathOf(landing);
  if (viewed) {
    const host = new URL(landing).host;
    const path = normalizeServerRelative(viewed);
    const byPath = `${fileApi(path, owningWeb(path, pageWebOf(page), host))}/$value`;
    const res = await client.getBinary(`https://${host}${byPath}`);
    return { res, name: splitParentLeaf(path).leaf };
  }
  const context = parseViewerPage(page);
  if (!context || !isSingleFileView(landing, url)) throw notAFile;
  const fileId = sourcedocOf(landing) ?? context.fileId;
  const res = await client.getBinary(
    `${context.webAbsoluteUrl}/_api/web/GetFileById('${fileId}')/$value`,
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
