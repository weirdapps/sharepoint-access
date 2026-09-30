import { describe, expect, it, vi } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { runLs } from '../src/commands/ls';
import { runLibraries } from '../src/commands/libraries';
import { runSearch } from '../src/commands/search';
import { runGet } from '../src/commands/get';
import { runPage } from '../src/commands/page';
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

  // Not by path: the web that owns a file cannot be read off its URL (a subsite is one segment
  // deeper), and a plain GET of a document URL already returns the file.
  it('fetches a direct file URL as is, and names it from its path', async () => {
    const getBinary = vi.fn().mockResolvedValue(bin());
    const url =
      'https://x.sharepoint.com/sites/hr/jobs/Lists/Roles/Attachments/3/Role%20profile.docx';
    const r = await runGet(binReader(getBinary), url);
    expect(getBinary).toHaveBeenCalledTimes(1);
    expect(getBinary.mock.calls[0][0]).toBe(url);
    expect(r.filename).toBe('Role profile.docx');
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

describe('runGet with sharing and viewer links', () => {
  const GUID = '0f8fad5b-d9cb-469f-a165-70867728950e';
  const file = (body = 'PK', filename?: string) => ({
    bytes: Buffer.from(body),
    contentType: 'application/octet-stream',
    ...(filename ? { filename } : {}),
  });
  const html = (body: string) => ({
    bytes: Buffer.from(body),
    contentType: 'text/html; charset=utf-8',
  });
  const viewer = html(
    `<html><script>var c = {"FileId":"${GUID}","FileName":"Budget.xlsx","webAbsoluteUrl":"https://x-my.sharepoint.com/personal/ann"};</script></html>`,
  );

  it('resolves an r sharing link through the viewer page it opens, subsite included', async () => {
    const subsiteViewer = html(
      `<script>var c = {"FileId":"${GUID}","webAbsoluteUrl":"https://x.sharepoint.com/sites/team/sub"};</script>`,
    );
    const getBinary = vi.fn().mockResolvedValueOnce(subsiteViewer).mockResolvedValueOnce(file());
    const url = 'https://x.sharepoint.com/:b:/r/sites/team/sub/Shared%20Documents/r.pdf?e=1';
    const r = await runGet(binReader(getBinary), url);
    expect(getBinary.mock.calls[0][0]).toBe(url);
    expect(getBinary.mock.calls[1][0]).toBe(
      `https://x.sharepoint.com/sites/team/sub/_api/web/GetFileById('${GUID}')/$value`,
    );
    expect(r.filename).toBe('r.pdf');
  });

  it('downloads the file a viewer page names, from the web that owns it', async () => {
    const getBinary = vi.fn().mockResolvedValueOnce(viewer).mockResolvedValueOnce(file('PK'));
    const r = await runGet(
      binReader(getBinary),
      'https://x-my.sharepoint.com/:x:/g/personal/ann/EQabc',
    );
    expect(getBinary).toHaveBeenCalledTimes(2);
    expect(getBinary.mock.calls[1][0]).toBe(
      `https://x-my.sharepoint.com/personal/ann/_api/web/GetFileById('${GUID}')/$value`,
    );
    expect(r.filename).toBe('Budget.xlsx');
    expect(r.size).toBe(2);
  });

  it('returns the first answer when it is already the file', async () => {
    const getBinary = vi.fn().mockResolvedValue(file('%PDF', 'r.pdf'));
    const r = await runGet(binReader(getBinary), 'https://x.sharepoint.com/:b:/g/sites/team/EQpdf');
    expect(getBinary).toHaveBeenCalledTimes(1);
    expect(r.filename).toBe('r.pdf');
  });

  it('fails not_a_file for a web page that names no file, and writes nothing', async () => {
    const out = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'sp-get-')), 'x.bin');
    const getBinary = vi.fn().mockResolvedValue(html('<html><title>Sign in</title></html>'));
    await expect(
      runGet(binReader(getBinary), 'https://x.sharepoint.com/:x:/g/sites/team/EQabc', out),
    ).rejects.toMatchObject({ code: 'NOT_A_FILE' });
    expect(fs.existsSync(out)).toBe(false);
  });

  // A OneDrive view carries FileIds of the files it lists; fetching one would store the wrong file.
  it('fails not_a_file for a link that is not a file, even when its page names one', async () => {
    const getBinary = vi.fn().mockResolvedValue(viewer);
    await expect(
      runGet(
        binReader(getBinary),
        'https://x-my.sharepoint.com/personal/ann/_layouts/15/onedrive.aspx',
      ),
    ).rejects.toMatchObject({ code: 'NOT_A_FILE' });
    expect(getBinary).toHaveBeenCalledTimes(1);
  });

  it('returns a link that is not a file when it answers with bytes', async () => {
    const getBinary = vi.fn().mockResolvedValue(file('IMG1'));
    const r = await runGet(
      binReader(getBinary),
      'https://x.sharepoint.com/sites/team/Assets/logo.png',
    );
    expect(r.size).toBe(4);
  });

  it('refuses an intranet page without fetching it', async () => {
    const getBinary = vi.fn();
    await expect(
      runGet(binReader(getBinary), 'https://x.sharepoint.com/sites/news/SitePages/Launch.aspx'),
    ).rejects.toMatchObject({ code: 'NOT_A_FILE' });
    expect(getBinary).not.toHaveBeenCalled();
  });
});

describe('runPage', () => {
  const PAGE = 'https://x.sharepoint.com/:u:/r/sites/news/SitePages/Launch.aspx?e=1';

  it('reads the page fields from the web that owns the page', async () => {
    const getJson = vi
      .fn()
      .mockResolvedValue({ Title: 'Launch', CanvasContent1: '<div>Hello</div>', WikiField: null });
    const r = await runPage(reader(getJson), PAGE);
    const url = getJson.mock.calls[0][0] as string;
    expect(url.startsWith('https://x.sharepoint.com/sites/news/_api/web/')).toBe(true);
    expect(url).toContain('/ListItemAllFields?$select=Title,CanvasContent1,WikiField');
    expect(r).toEqual({
      source: PAGE,
      path: '/sites/news/SitePages/Launch.aspx',
      title: 'Launch',
      html: '<div>Hello</div>',
    });
  });

  it('reads a page in a subsite from that subsite', async () => {
    const getJson = vi.fn().mockResolvedValue({ Title: 'T', CanvasContent1: '<p>x</p>' });
    await runPage(reader(getJson), 'https://x.sharepoint.com/sites/hr/news/SitePages/3684.aspx');
    const url = getJson.mock.calls[0][0] as string;
    expect(url.startsWith('https://x.sharepoint.com/sites/hr/news/_api/web/')).toBe(true);
  });

  it('falls back to the wiki field of an older page', async () => {
    const getJson = vi
      .fn()
      .mockResolvedValue({ Title: 'Old', CanvasContent1: null, WikiField: '<p>Wiki text</p>' });
    expect((await runPage(reader(getJson), PAGE)).html).toBe('<p>Wiki text</p>');
  });

  it('rejects a link that is not a page', async () => {
    const getJson = vi.fn();
    await expect(
      runPage(reader(getJson), 'https://x.sharepoint.com/sites/t/Shared%20Documents/a.docx'),
    ).rejects.toBeInstanceOf(PathError);
    expect(getJson).not.toHaveBeenCalled();
  });
});
