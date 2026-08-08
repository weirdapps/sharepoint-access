import { describe, expect, it, vi } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { runMkdir } from '../src/commands/mkdir';
import { runPut } from '../src/commands/put';
import { PathError } from '../src/sharepoint/paths';

describe('runMkdir', () => {
  it('posts to folders/addUsingPath', async () => {
    const postJson = vi.fn().mockResolvedValue({});
    await runMkdir({ postJson } as never, '/a/New');
    const url = postJson.mock.calls[0][0] as string;
    expect(url).toContain('/_api/web/folders/addUsingPath');
    expect(url).toContain('DecodedUrl=');
  });

  it('percent-encodes a Greek folder name', async () => {
    const postJson = vi.fn().mockResolvedValue({});
    await runMkdir({ postJson } as never, '/Έγγραφα/Νέος');
    expect(postJson.mock.calls[0][0] as string).toContain('%CE%9D');
  });

  it('rejects the root', async () => {
    const postJson = vi.fn();
    await expect(runMkdir({ postJson } as never, '/')).rejects.toThrowError(PathError);
    expect(postJson).not.toHaveBeenCalled();
  });

  it('rejects a traversal path before making a request', async () => {
    const postJson = vi.fn();
    await expect(runMkdir({ postJson } as never, '/a/../../x')).rejects.toThrowError(PathError);
    expect(postJson).not.toHaveBeenCalled();
  });

  it('returns the normalized path it created', async () => {
    const postJson = vi.fn().mockResolvedValue({});
    expect((await runMkdir({ postJson } as never, '/a//New/')).serverRelativeUrl).toBe('/a/New');
  });
});

describe('runPut', () => {
  function tmpFile(name: string, contents = 'hello'): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spp-'));
    const p = path.join(dir, name);
    fs.writeFileSync(p, contents);
    return p;
  }

  const client = () => ({
    postBinary: vi.fn().mockResolvedValue({}),
    postJson: vi.fn().mockResolvedValue({}),
  });

  it('derives the leaf from the local basename', async () => {
    const c = client();
    const local = tmpFile('report.docx');
    await runPut(c as never, local, '/a', true);
    expect(c.postBinary.mock.calls[0][0] as string).toContain('report.docx');
  });

  it('handles a Greek local filename', async () => {
    const c = client();
    const local = tmpFile('έκθεση.docx');
    const r = await runPut(c as never, local, '/a', true);
    expect(r.serverRelativeUrl).toBe('/a/έκθεση.docx');
    expect(c.postBinary.mock.calls[0][0] as string).toContain('%CE%AD');
  });

  it('uploads the actual file contents', async () => {
    const c = client();
    const local = tmpFile('f.txt', 'payload-bytes');
    await runPut(c as never, local, '/a', true);
    expect((c.postBinary.mock.calls[0][1] as Buffer).toString()).toBe('payload-bytes');
  });

  it('propagates overwrite', async () => {
    const c = client();
    await runPut(c as never, tmpFile('f.txt'), '/a', false);
    expect(c.postBinary.mock.calls[0][0] as string).toContain('overwrite=false');
  });

  it('rejects a missing local file with an IO error', async () => {
    const c = client();
    await expect(runPut(c as never, '/definitely/not/here.txt', '/a', true)).rejects.toMatchObject({
      code: 'IO',
    });
    expect(c.postBinary).not.toHaveBeenCalled();
  });
});
