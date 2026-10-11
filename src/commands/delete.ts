import type { SharepointClient } from '../http/client';
import { fileApi, normalizeServerRelative } from '../sharepoint/paths';

type Deleter = Pick<SharepointClient, 'deleteFile'>;

export async function runDelete(
  client: Deleter,
  path: string,
  site?: string,
): Promise<{ serverRelativeUrl: string; deleted: boolean }> {
  const normalized = normalizeServerRelative(path);

  await client.deleteFile(fileApi(normalized, site));

  return {
    serverRelativeUrl: normalized,
    deleted: true,
  };
}
