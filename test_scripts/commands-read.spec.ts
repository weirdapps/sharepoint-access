import { describe, expect, it, vi } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { runLs } from '../src/commands/ls';
import { runLibraries } from '../src/commands/libraries';
import { runSearch } from '../src/commands/search';
import { runGet } from '../src/commands/get';
import { PathError } from '../src/sharepoint/paths';

const reader = (getJson: unknown) => ({ getJson }) as never;
const binReader = (getBinary: unknown) => ({ getBinary }) as never;

describe('runLs', () => {
  it('expands Folders and Files in one request', async () => {
    const getJson = vi.fn().mockResolvedValue({ Folders: [], Files: [] });
    await runLs(reader(getJson), '/Έγγραφα');
    expect(getJson).toHaveBeenCalledTimes(1);
    const url = getJson.mock.calls[0][0] as string;
    expect(url).toContain('GetFolderByServerRelativePath');
    expect(url).toContain('$expand=Folders,Files');
  });

  it('maps folders and files into a stable shape', async () => {
    const getJson = vi.fn().mockResolvedValue({
      Folders: [
        { Name: 'Sub', ServerRelativeUrl: '/a/Sub', TimeLastModified: '2026-01-01T00:00:00Z' },
      ],
      Files: [
        {
          Name: 'α.docx',
          ServerRelativeUrl: '/a/α.docx',
          Length: '42',
          TimeLastModified: '2026-01-02T00:00:00Z',
        },
      ],
    });
    const r = await runLs(reader(getJson), '/a');
    expect(r.folders).toEqual([
      { name: 'Sub', serverRelativeUrl: '/a/Sub', modified: '2026-01-01T00:00:00Z' },
    ]);
    expect(r.files).toEqual([
      {
        name: 'α.docx',
        serverRelativeUrl: '/a/α.docx',
        size: 42,
        modified: '2026-01-02T00:00:00Z',
      },
    ]);
  });

  it('coerces the string Length SharePoint returns into a number', async () => {
    const getJson = vi.fn().mockResolvedValue({
      Folders: [],
      Files: [{ Name: 'f', ServerRelativeUrl: '/f', Length: '1048576' }],
    });
    const r = await runLs(reader(getJson), '/');
    expect(r.files[0].size).toBe(1048576);
  });

  it('omits size when Length is absent rather than reporting NaN', async () => {
    const getJson = vi
      .fn()
      .mockResolvedValue({ Folders: [], Files: [{ Name: 'f', ServerRelativeUrl: '/f' }] });
    const r = await runLs(reader(getJson), '/');
    expect(r.files[0].size).toBeUndefined();
    expect('size' in r.files[0]).toBe(false);
  });

  it('tolerates a response with neither collection present', async () => {
    const r = await runLs(reader(vi.fn().mockResolvedValue({})), '/a');
    expect(r.folders).toEqual([]);
    expect(r.files).toEqual([]);
  });

  it('returns the normalized path it queried', async () => {
    const r = await runLs(reader(vi.fn().mockResolvedValue({})), '/a//b/');
    expect(r.path).toBe('/a/b');
  });

  it('rejects a traversal path before making a request', async () => {
    const getJson = vi.fn();
    await expect(runLs(reader(getJson), '/a/../../etc')).rejects.toThrowError(PathError);
    expect(getJson).not.toHaveBeenCalled();
  });
});

describe('runLibraries', () => {
  it('filters to visible document libraries', async () => {
    const getJson = vi.fn().mockResolvedValue({ value: [] });
    await runLibraries(reader(getJson));
    const url = getJson.mock.calls[0][0] as string;
    expect(url).toContain('BaseTemplate%20eq%20101');
    expect(url).toContain('Hidden%20eq%20false');
  });

  it('maps the root folder URL out of the expansion', async () => {
    const getJson = vi.fn().mockResolvedValue({
      value: [{ Title: 'Έγγραφα', Id: 'guid-1', RootFolder: { ServerRelativeUrl: '/Έγγραφα' } }],
    });
    const r = await runLibraries(reader(getJson));
    expect(r.libraries).toEqual([
      { title: 'Έγγραφα', id: 'guid-1', serverRelativeUrl: '/Έγγραφα' },
    ]);
  });

  it('tolerates a library with no expanded root folder', async () => {
    const getJson = vi.fn().mockResolvedValue({ value: [{ Title: 'T', Id: 'g' }] });
    expect((await runLibraries(reader(getJson))).libraries[0].serverRelativeUrl).toBe('');
  });
});

describe('runSearch', () => {
  const empty = { PrimaryQueryResult: { RelevantResults: { Table: { Rows: [] } } } };

  it('quotes the query and passes rowlimit', async () => {
    const getJson = vi.fn().mockResolvedValue(empty);
    await runSearch(reader(getJson), 'budget', 5);
    const url = getJson.mock.calls[0][0] as string;
    expect(url).toContain("querytext='budget'");
    expect(url).toContain('rowlimit=5');
  });

  it('escapes an apostrophe in the query rather than breaking the literal', async () => {
    const getJson = vi.fn().mockResolvedValue(empty);
    await runSearch(reader(getJson), "O'Brien", 5);
    expect(getJson.mock.calls[0][0] as string).toContain("O''Brien");
  });

  it('flattens the Cells key/value rows into objects', async () => {
    const getJson = vi.fn().mockResolvedValue({
      PrimaryQueryResult: {
        RelevantResults: {
          TotalRows: 1,
          Table: {
            Rows: [
              {
                Cells: [
                  { Key: 'Title', Value: 'T' },
                  { Key: 'Path', Value: 'https://x/a' },
                ],
              },
            ],
          },
        },
      },
    });
    const r = await runSearch(reader(getJson), 'q', 1);
    expect(r.results).toEqual([{ Title: 'T', Path: 'https://x/a' }]);
    expect(r.totalRows).toBe(1);
  });

  it('drops cells with a null value rather than emitting null', async () => {
    const getJson = vi.fn().mockResolvedValue({
      PrimaryQueryResult: {
        RelevantResults: {
          Table: {
            Rows: [
              {
                Cells: [
                  { Key: 'Title', Value: null },
                  { Key: 'Path', Value: 'p' },
                ],
              },
            ],
          },
        },
      },
    });
    expect((await runSearch(reader(getJson), 'q', 1)).results).toEqual([{ Path: 'p' }]);
  });

  it('rejects a non-positive rowlimit', async () => {
    await expect(runSearch(reader(vi.fn()), 'q', 0)).rejects.toThrowError(/rows/);
  });

  it('rejects an absurd rowlimit rather than hammering the search service', async () => {
    await expect(runSearch(reader(vi.fn()), 'q', 100000)).rejects.toThrowError(/rows/);
  });

  it('rejects an empty query', async () => {
    await expect(runSearch(reader(vi.fn()), '   ', 5)).rejects.toThrowError(/query/);
  });
});

describe('runGet', () => {
  const bin = (body = 'x') => ({ bytes: Buffer.from(body), contentType: 'text/plain' });

  it('appends /$value to the file accessor', async () => {
    const getBinary = vi.fn().mockResolvedValue(bin());
    await runGet(binReader(getBinary), '/a/b.txt');
    expect(getBinary.mock.calls[0][0] as string).toMatch(/GetFileByServerRelativePath.*\/\$value$/);
  });

  it('passes an absolute URL straight through for host checking downstream', async () => {
    const getBinary = vi.fn().mockResolvedValue(bin());
    await runGet(binReader(getBinary), 'https://x.sharepoint.com/a.docx');
    expect(getBinary.mock.calls[0][0]).toBe('https://x.sharepoint.com/a.docx');
  });

  it('writes bytes to disk when an out path is given', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spg-'));
    const out = path.join(dir, 'nested', 'f.txt');
    const getBinary = vi.fn().mockResolvedValue(bin('hello'));
    const r = await runGet(binReader(getBinary), '/a/f.txt', out);
    expect(fs.readFileSync(out, 'utf8')).toBe('hello');
    expect(r.outPath).toBe(path.resolve(out));
    expect(r.size).toBe(5);
  });

  it('derives the filename from the path when the server sends none', async () => {
    const getBinary = vi.fn().mockResolvedValue(bin());
    expect((await runGet(binReader(getBinary), '/a/α.docx')).filename).toBe('α.docx');
  });

  it('prefers the server-supplied filename', async () => {
    const getBinary = vi.fn().mockResolvedValue({ ...bin(), filename: 'server-name.docx' });
    expect((await runGet(binReader(getBinary), '/a/local.docx')).filename).toBe('server-name.docx');
  });
});
