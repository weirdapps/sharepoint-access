// src/commands/page.ts
//
// An intranet page's title and content. A plain GET of a modern page returns a script shell of
// about 900 KB that renders the text in the browser; the text itself is the page's list item
// field CanvasContent1 (WikiField on an older wiki page), read through REST from the web that
// owns the page.

import type { SharepointClient } from '../http/client';
import { serverRelativeFromUrl } from '../sharepoint/links';
import { PathError, fileApi, normalizeServerRelative } from '../sharepoint/paths';

export interface PageResult {
  source: string;
  path: string;
  title: string;
  /** The page's content as HTML: CanvasContent1, else WikiField, else empty. */
  html: string;
}

interface PageFields {
  Title?: string | null;
  CanvasContent1?: string | null;
  WikiField?: string | null;
}

type Reader = Pick<SharepointClient, 'getJson'>;

function isAbsoluteUrl(s: string): boolean {
  return /^[a-z][a-z0-9+.-]*:/i.test(s);
}

export async function runPage(client: Reader, url: string, site?: string): Promise<PageResult> {
  const absolute = isAbsoluteUrl(url);
  const path = normalizeServerRelative(absolute ? serverRelativeFromUrl(url) : url);
  if (!/\.aspx$/i.test(path)) {
    throw new PathError(`page: "${url}" is not a page (.aspx)`);
  }
  const api = `${fileApi(path, site)}/ListItemAllFields?$select=Title,CanvasContent1,WikiField`;
  const fields = await client.getJson<PageFields>(
    absolute ? `https://${new URL(url).host}${api}` : api,
  );
  return {
    source: url,
    path,
    title: fields.Title ?? '',
    html: fields.CanvasContent1 || fields.WikiField || '',
  };
}
