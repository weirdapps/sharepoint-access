// src/commands/ls.ts

import type { SharepointClient } from '../http/client';
import type { Entry, SpFile, SpFolder, SpFolderListing } from '../http/types';
import { folderApi, normalizeServerRelative } from '../sharepoint/paths';

export interface LsResult {
  path: string;
  folders: Entry[];
  files: Entry[];
}

type Reader = Pick<SharepointClient, 'getJson'>;

const SELECT = [
  'Name',
  'ServerRelativeUrl',
  'Folders/Name',
  'Folders/ServerRelativeUrl',
  'Folders/TimeLastModified',
  'Files/Name',
  'Files/ServerRelativeUrl',
  'Files/Length',
  'Files/TimeLastModified',
].join(',');

/** SharePoint returns Length as a decimal string; callers want a number. */
function toSize(v: SpFile['Length']): number | undefined {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string' || v.length === 0) return undefined;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

function mapFolder(f: SpFolder): Entry {
  return {
    name: f.Name ?? '',
    serverRelativeUrl: f.ServerRelativeUrl ?? '',
    ...(f.TimeLastModified ? { modified: f.TimeLastModified } : {}),
  };
}

function mapFile(f: SpFile): Entry {
  const size = toSize(f.Length);
  return {
    name: f.Name ?? '',
    serverRelativeUrl: f.ServerRelativeUrl ?? '',
    ...(size !== undefined ? { size } : {}),
    ...(f.TimeLastModified ? { modified: f.TimeLastModified } : {}),
  };
}

export async function runLs(client: Reader, path: string): Promise<LsResult> {
  const norm = normalizeServerRelative(path);
  const url = `${folderApi(norm)}?$expand=Folders,Files&$select=${encodeURIComponent(SELECT)}`;
  const body = await client.getJson<SpFolderListing>(url);
  return {
    path: norm,
    folders: (body.Folders ?? []).map(mapFolder),
    files: (body.Files ?? []).map(mapFile),
  };
}
