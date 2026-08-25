// test_scripts/web-scoping.spec.ts
//
// Regression tests for the sub-site scoping bug found during live testing.
//
// SharePoint's `_api` calls run against a specific WEB. Asking the host's root
// web about a path it does not own returns 200 with EMPTY collections, not a
// 404, so getting this wrong yields a confidently empty listing rather than an
// error. Verified on the live tenant 2026-08-08:
//
//   root web + /personal/<user>/Documents -> 200 {"Folders":[]}
//   /personal/<user> web + same path      -> 200 with 6 folders

import { describe, expect, it, vi } from 'vitest';

import {
  webPrefixFor,
  normalizeWeb,
  folderApi,
  fileApi,
  foldersAddApi,
  listsApi,
} from '../src/sharepoint/paths';
import { runLs } from '../src/commands/ls';
import { runMkdir } from '../src/commands/mkdir';
import { uploadFile } from '../src/sharepoint/upload';

describe('webPrefixFor', () => {
  it.each([
    ['/personal/user_contoso_com/Documents', '/personal/user_contoso_com'],
    ['/personal/user_x/Documents/Deep/Nested/File.docx', '/personal/user_x'],
    ['/sites/div991secb/Shared Documents', '/sites/div991secb'],
    ['/teams/someteam/Docs', '/teams/someteam'],
    ['/DRPdocuments', ''],
    ['/DRPdocuments/Forms/file.docx', ''],
    ['/', ''],
  ])('maps %s to web %s', (path, expected) => {
    expect(webPrefixFor(path)).toBe(expected);
  });

  it('is case-insensitive on the web root segment', () => {
    expect(webPrefixFor('/Sites/Foo/Docs')).toBe('/Sites/Foo');
  });

  it('does not treat a single segment as a web', () => {
    // "/sites" alone is a library named sites, not a site collection.
    expect(webPrefixFor('/sites')).toBe('');
  });
});

describe('normalizeWeb', () => {
  it('returns undefined when no override is given, so the heuristic applies', () => {
    expect(normalizeWeb(undefined)).toBeUndefined();
  });

  it('maps an explicit empty string or slash to the root web', () => {
    expect(normalizeWeb('')).toBe('');
    expect(normalizeWeb('/')).toBe('');
  });

  it('normalizes an explicit site path', () => {
    expect(normalizeWeb('/sites/a/')).toBe('/sites/a');
    expect(normalizeWeb('sites/a')).toBe('/sites/a');
  });
});

describe('API builders are web-scoped', () => {
  it('prefixes the personal web on folderApi', () => {
    expect(folderApi('/personal/u/Documents')).toMatch(
      /^\/personal\/u\/_api\/web\/GetFolderByServerRelativePath/,
    );
  });

  it('prefixes the site web on fileApi', () => {
    expect(fileApi('/sites/s/Shared Documents/a.docx')).toMatch(
      /^\/sites\/s\/_api\/web\/GetFileByServerRelativePath/,
    );
  });

  it('leaves a root-web path unprefixed', () => {
    expect(folderApi('/DRPdocuments')).toMatch(/^\/_api\/web\//);
  });

  it('honours an explicit site override for a deeply nested web', () => {
    expect(folderApi('/sites/a/sub/Docs', '/sites/a/sub')).toMatch(/^\/sites\/a\/sub\/_api\/web\//);
  });

  it('lets an explicit empty override force the root web', () => {
    expect(folderApi('/sites/a/Docs', '')).toMatch(/^\/_api\/web\//);
  });

  it('web-scopes folder creation', () => {
    expect(foldersAddApi('/personal/u/Documents/New')).toMatch(
      /^\/personal\/u\/_api\/web\/folders\/addUsingPath/,
    );
  });

  it('web-scopes the lists endpoint', () => {
    expect(listsApi('/sites/s')).toBe('/sites/s/_api/web/lists');
    expect(listsApi(undefined)).toBe('/_api/web/lists');
  });
});

describe('commands use the scoped endpoint', () => {
  it('runLs targets the personal web', async () => {
    const getJson = vi.fn().mockResolvedValue({ Exists: true, Folders: [], Files: [] });
    await runLs({ getJson } as never, '/personal/u/Documents');
    expect(getJson.mock.calls[0][0] as string).toMatch(/^\/personal\/u\/_api\/web\//);
  });

  it('runLs raises NOT_FOUND when the folder does not exist', async () => {
    // Without this, a phantom path reads as an empty folder.
    const getJson = vi.fn().mockResolvedValue({ Exists: false });
    await expect(runLs({ getJson } as never, '/nope')).rejects.toMatchObject({
      code: 'NOT_FOUND',
    });
  });

  it('runLs asks for Exists so the phantom case is detectable at all', async () => {
    const getJson = vi.fn().mockResolvedValue({ Exists: true, Folders: [], Files: [] });
    await runLs({ getJson } as never, '/DRPdocuments');
    expect(decodeURIComponent(getJson.mock.calls[0][0] as string)).toContain('Exists');
  });

  it('runMkdir targets the site web', async () => {
    const postJson = vi.fn().mockResolvedValue({});
    await runMkdir({ postJson } as never, '/sites/s/Docs/New');
    expect(postJson.mock.calls[0][0] as string).toMatch(/^\/sites\/s\/_api\/web\/folders\//);
  });

  it('uploadFile targets the personal web for every request in a chunked session', async () => {
    const c = {
      postBinary: vi.fn().mockResolvedValue({}),
      postJson: vi.fn().mockResolvedValue({}),
    };
    await uploadFile(
      c as never,
      Buffer.alloc(25 * 1024 * 1024),
      '/personal/u/Documents',
      'big.bin',
      true,
    );
    for (const call of c.postBinary.mock.calls) {
      expect(call[0] as string).toMatch(/^\/personal\/u\/_api\/web\//);
    }
  });
});
