// src/commands/libraries.ts

import type { SharepointClient } from '../http/client';
import type { SpListsResponse } from '../http/types';

export interface Library {
  title: string;
  id: string;
  serverRelativeUrl: string;
}

export interface LibrariesResult {
  libraries: Library[];
}

type Reader = Pick<SharepointClient, 'getJson'>;

/** BaseTemplate 101 is the document-library template. */
const FILTER = 'BaseTemplate eq 101 and Hidden eq false';

export async function runLibraries(client: Reader): Promise<LibrariesResult> {
  const url =
    `/_api/web/lists?$filter=${encodeURIComponent(FILTER)}` +
    `&$select=${encodeURIComponent('Title,Id,RootFolder/ServerRelativeUrl')}` +
    `&$expand=RootFolder`;
  const body = await client.getJson<SpListsResponse>(url);
  return {
    libraries: (body.value ?? []).map((l) => ({
      title: l.Title ?? '',
      id: l.Id ?? '',
      serverRelativeUrl: l.RootFolder?.ServerRelativeUrl ?? '',
    })),
  };
}
