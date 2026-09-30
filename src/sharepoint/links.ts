// src/sharepoint/links.ts
//
// What a SharePoint URL found in mail points at, and how to reach its content.
//
// A sharing link (/:x:/g/...) or a viewer URL answers a plain GET with the browser page, not the
// file. That page's context names the file (FileId) and the web that owns it (webAbsoluteUrl),
// which is all GetFileById needs. The owning web cannot be read off a URL: a subsite is one
// segment deeper than a site, and nothing in the path says which segment is a library.
// The /_api/v2.0/shares endpoint would resolve any link, but it refuses a cookie session (403).

export type LinkKind = 'file' | 'page' | 'other';

export interface ViewerContext {
  fileId: string;
  webAbsoluteUrl: string;
  fileName?: string;
}

/** Sharing-link letters that stand for a file: Word, Excel, PowerPoint, PDF, text, OneNote. */
const FILE_LETTERS = new Set(['w', 'x', 'p', 'b', 't', 'o']);

/** "/:<letter>:/<form>/<rest>", where form r carries the server-relative path. */
const SHARING_RE = /^\/:([a-z]):\/([a-z])(\/.*)?$/i;

const PAGE_RE = /\/SitePages\/[^/]+\.aspx$/i;

/** Office Online viewers: the file is named in the query (sourcedoc) and in the page context. */
const VIEWER_RE = /\/_layouts\/15\/(Doc|WopiFrame2?|xlviewer|PowerPoint)\.aspx$/i;

const DOCUMENT_RE =
  /\.(docx?|docm|dotx|xlsx?|xlsm|xlsb|pptx?|pptm|ppsx|pdf|txt|csv|md|rtf|odt|ods|odp|msg|eml|zip)$/i;

const GUID = '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})';
// Quoted JSON keys only: the page's scripts also say fileId:w.FileId, which is code, not data.
const FILE_ID_RE = new RegExp(`"FileId"\\s*:\\s*"\\{?${GUID}\\}?"`, 'i');
const SOURCEDOC_RE = new RegExp(`sourcedoc=(?:%7B|\\{)?${GUID}`, 'i');
const WEB_RE = /"webAbsoluteUrl"\s*:\s*"((?:[^"\\]|\\.)*)"/;
const NAME_RE = /"FileName"\s*:\s*"((?:[^"\\]|\\.)*)"/;

function pathOf(url: string): string {
  return decodeURIComponent(new URL(url).pathname);
}

/** The percent-decoded server-relative path of a URL, query dropped, "r" sharing prefix removed. */
export function serverRelativeFromUrl(url: string): string {
  const path = pathOf(url);
  const m = SHARING_RE.exec(path);
  if (m && m[2].toLowerCase() === 'r' && m[3]) return m[3];
  return path;
}

/** Whether a link is a file, an intranet page, or neither (home, views, settings, folders). */
export function classifyLink(url: string): LinkKind {
  let path: string;
  try {
    path = pathOf(url);
  } catch {
    return 'other';
  }
  const m = SHARING_RE.exec(path);
  if (m) {
    const letter = m[1].toLowerCase();
    if (letter === 'u' && m[2].toLowerCase() === 'r' && PAGE_RE.test(m[3] ?? '')) return 'page';
    return FILE_LETTERS.has(letter) ? 'file' : 'other';
  }
  if (VIEWER_RE.test(path)) return 'file';
  if (PAGE_RE.test(path)) return 'page';
  if (DOCUMENT_RE.test(path)) return 'file';
  return 'other';
}

/** The server-relative path of a URL that names a document (a document URL, or an r sharing
 * link to one), or null when only the page behind the link can say. */
export function filePathOf(url: string): string | null {
  let path: string;
  try {
    path = serverRelativeFromUrl(url);
  } catch {
    return null;
  }
  return DOCUMENT_RE.test(path) ? path : null;
}

/** The server-relative path a list view names in its id parameter: a OneDrive or library view
 * opened on one file, where a PDF or text sharing link lands. Null when there is none. */
export function viewedPathOf(url: string): string | null {
  try {
    const id = new URL(url).searchParams.get('id');
    return id && id.startsWith('/') ? id : null;
  } catch {
    return null;
  }
}

/** The file a viewer URL opens by ID (sourcedoc={guid}), lower-cased, or null. */
export function sourcedocOf(url: string): string | null {
  try {
    const m = new RegExp(`^\\{?${GUID}\\}?$`, 'i').exec(
      new URL(url).searchParams.get('sourcedoc') ?? '',
    );
    return m ? m[1].toLowerCase() : null;
  } catch {
    return null;
  }
}

/** Whether the page at `landing` shows one file: the link itself, an Office viewer, or a
 * document URL. A list view does not: its context names every file it lists. */
export function isSingleFileView(landing: string, link: string): boolean {
  try {
    const path = pathOf(landing);
    return path === pathOf(link) || VIEWER_RE.test(path) || DOCUMENT_RE.test(path);
  } catch {
    return false;
  }
}

function jsonString(raw: string): string | undefined {
  try {
    return JSON.parse(`"${raw}"`) as string;
  } catch {
    return undefined;
  }
}

/** The web a page belongs to (its context's webAbsoluteUrl), without a trailing slash. */
export function pageWebOf(html: string): string | undefined {
  const raw = WEB_RE.exec(html)?.[1];
  const web = raw === undefined ? undefined : jsonString(raw);
  return web && /^https:\/\//i.test(web) ? web.replace(/\/+$/, '') : undefined;
}

/** The file a viewer page shows, from its context, or null when the page names no file. */
export function parseViewerPage(html: string): ViewerContext | null {
  const id = FILE_ID_RE.exec(html)?.[1] ?? SOURCEDOC_RE.exec(html)?.[1];
  const web = pageWebOf(html);
  if (!id || !web) return null;
  const nameRaw = NAME_RE.exec(html)?.[1];
  const fileName = nameRaw === undefined ? undefined : jsonString(nameRaw);
  return { fileId: id.toLowerCase(), webAbsoluteUrl: web, ...(fileName ? { fileName } : {}) };
}
