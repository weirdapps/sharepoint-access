// src/commands/ls.ts

import type { SharepointClient } from '../http/client';
import type { Entry, SpFile, SpFolder, SpFolderListing } from '../http/types';
import { CliError } from '../config/errors';
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
  // Exists turns a phantom path into an explicit error. The root web answers
  // 200 for paths it does not own, so without this a wrong path or wrong web
  // reads as a confidently empty folder.
  'Exists',
  'ItemCount',
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

export async function runLs(client: Reader, path: string, site?: string): Promise<LsResult> {
  const norm = normalizeServerRelative(path);
  const url = `${folderApi(norm, site)}?$expand=Folders,Files&$select=${encodeURIComponent(SELECT)}`;
  const body = await client.getJson<SpFolderListing>(url);
  if (body.Exists === false) {
    throw new CliError('NOT_FOUND', `folder does not exist: "${norm}"`);
  }
  return {
    path: norm,
    folders: (body.Folders ?? []).map(mapFolder),
    files: (body.Files ?? []).map(mapFile),
  };
}
