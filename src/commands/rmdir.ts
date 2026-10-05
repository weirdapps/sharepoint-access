import type { SharepointClient } from '../http/client';
import { folderApi, normalizeServerRelative } from '../sharepoint/paths';

type FolderRemover = Pick<SharepointClient, 'deleteFolder'>;

export async function runRmdir(
  client: FolderRemover,
  path: string,
  site?: string,
): Promise<{ serverRelativeUrl: string; deleted: boolean }> {
  const normalized = normalizeServerRelative(path);

  await client.deleteFolder(folderApi(normalized, site));

  return {
    serverRelativeUrl: normalized,
    deleted: true,
  };
}
