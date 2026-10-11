// src/commands/inventory.ts

import type { SharepointClient } from '../http/client';
import type { Entry } from '../http/types';
import { folderApi, normalizeServerRelative } from '../sharepoint/paths';
import { runLs, type LsResult } from './ls';

type Reader = Pick<SharepointClient, 'getJson'>;

export interface InventoryEntry {
  type: 'D' | 'F';
  path: string;
  size?: number;
  modified?: string;
}

export interface InventoryResult {
  path: string;
  entries: InventoryEntry[];
}

function addFolder(entries: InventoryEntry[], folder: Entry): void {
  entries.push({
    type: 'D',
    path: folder.serverRelativeUrl,
    ...(folder.modified ? { modified: folder.modified } : {}),
  });
}

function addFile(entries: InventoryEntry[], file: Entry): void {
  entries.push({
    type: 'F',
    path: file.serverRelativeUrl,
    ...(file.size !== undefined ? { size: file.size } : {}),
    ...(file.modified ? { modified: file.modified } : {}),
  });
}

export async function runInventory(
  client: Reader,
  path: string,
  site?: string,
): Promise<InventoryResult> {
  const root = normalizeServerRelative(path);
  const entries: InventoryEntry[] = [];

  // Fetch one directory, then recurse into all child directories concurrently.
  const first = await runLs(client, root, site);

  for (const folder of first.folders) addFolder(entries, folder);
  for (const file of first.files) addFile(entries, file);

  async function walk(folders: Entry[]): Promise<void> {
    await Promise.all(
      folders.map(async (folder) => {
        const listing: LsResult = await runLs(client, folder.serverRelativeUrl, site);

        for (const child of listing.folders) addFolder(entries, child);
        for (const file of listing.files) addFile(entries, file);

        await walk(listing.folders);
      }),
    );
  }

  await walk(first.folders);

  return { path: root, entries };
}
