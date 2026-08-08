import { describe, expect, it, vi } from 'vitest';

import { uploadFile, CHUNK_THRESHOLD_BYTES, CHUNK_SIZE_BYTES } from '../src/sharepoint/upload';

function mkClient() {
  return {
    postBinary: vi.fn().mockResolvedValue({ ServerRelativeUrl: '/a/f.bin' }),
    postJson: vi.fn().mockResolvedValue({}),
  };
}

const urlsOf = (c: ReturnType<typeof mkClient>) =>
  c.postBinary.mock.calls.map((x) => x[0] as string);

describe('uploadFile', () => {
  it('uses a single addUsingPath request below the threshold', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(1024), '/a', 'f.bin', true);
    expect(c.postBinary).toHaveBeenCalledTimes(1);
    const url = urlsOf(c)[0];
    expect(url).toContain('addUsingPath');
    expect(url).toContain('overwrite=true');
    expect(url).not.toContain('StartUpload');
  });

  it('propagates overwrite=false', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(16), '/a', 'f.bin', false);
    expect(urlsOf(c)[0]).toContain('overwrite=false');
  });

  it('never uses the legacy add(url=) form', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(16), '/a', 'f.bin', false);
    expect(urlsOf(c)[0]).not.toMatch(/\/add\(url=/);
  });

  it('percent-encodes a Greek leaf name', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(16), '/a', 'α.docx', true);
    expect(urlsOf(c)[0]).toContain('%CE%B1');
  });

  it('doubles an apostrophe in the leaf name', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(16), '/a', "O'B.docx", true);
    expect(urlsOf(c)[0]).toContain("O''B");
  });

  it('rejects a leaf containing a slash rather than creating a nested path', async () => {
    const c = mkClient();
    await expect(
      uploadFile(c as never, Buffer.alloc(16), '/a', '../evil.docx', true),
    ).rejects.toThrowError();
    expect(c.postBinary).not.toHaveBeenCalled();
  });

  it('returns the composed server-relative target', async () => {
    const c = mkClient();
    const r = await uploadFile(c as never, Buffer.alloc(16), '/a/b', 'f.bin', true);
    expect(r.serverRelativeUrl).toBe('/a/b/f.bin');
    expect(r.chunked).toBe(false);
    expect(r.size).toBe(16);
  });

  it('chunks at exactly the threshold', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_THRESHOLD_BYTES), '/a', 'big.bin', true);
    const urls = urlsOf(c);
    expect(urls.some((u) => u.includes('StartUpload'))).toBe(true);
    expect(urls.some((u) => u.includes('FinishUpload'))).toBe(true);
  });

  it('does not chunk one byte below the threshold', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_THRESHOLD_BYTES - 1), '/a', 'f.bin', true);
    expect(c.postBinary).toHaveBeenCalledTimes(1);
  });

  it('creates the empty file first, then starts the upload session', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_THRESHOLD_BYTES), '/a', 'big.bin', true);
    const urls = urlsOf(c);
    expect(urls[0]).toContain('addUsingPath');
    expect(urls[1]).toContain('StartUpload');
    // The placeholder must be empty, not the payload.
    expect((c.postBinary.mock.calls[0][1] as Buffer).length).toBe(0);
  });

  it('splits an exact multiple into Start plus Continues plus Finish with no empty final chunk', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 3), '/a', 'big.bin', true);
    const urls = urlsOf(c);
    expect(urls.filter((u) => u.includes('StartUpload')).length).toBe(1);
    expect(urls.filter((u) => u.includes('ContinueUpload')).length).toBe(1);
    expect(urls.filter((u) => u.includes('FinishUpload')).length).toBe(1);
    // Every chunk carries bytes.
    for (const call of c.postBinary.mock.calls.slice(1)) {
      expect((call[1] as Buffer).length).toBeGreaterThan(0);
    }
  });

  it('sends correct fileOffset values', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 2 + 5), '/a', 'big.bin', true);
    const urls = urlsOf(c);
    expect(urls.find((u) => u.includes('ContinueUpload'))).toContain(
      `fileOffset=${CHUNK_SIZE_BYTES}`,
    );
    expect(urls.find((u) => u.includes('FinishUpload'))).toContain(
      `fileOffset=${CHUNK_SIZE_BYTES * 2}`,
    );
  });

  it('sends every byte exactly once across the session', async () => {
    const c = mkClient();
    const size = CHUNK_SIZE_BYTES * 2 + 12345;
    await uploadFile(c as never, Buffer.alloc(size), '/a', 'big.bin', true);
    const sent = c.postBinary.mock.calls
      .slice(1)
      .reduce((acc, call) => acc + (call[1] as Buffer).length, 0);
    expect(sent).toBe(size);
  });

  it('uses one uploadId GUID for the whole session', async () => {
    const c = mkClient();
    await uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 2 + 1), '/a', 'big.bin', true);
    const ids = urlsOf(c)
      .map((u) => u.match(/uploadId=guid'([^']+)'/)?.[1])
      .filter(Boolean);
    expect(ids.length).toBeGreaterThan(1);
    expect(new Set(ids).size).toBe(1);
  });

  it('deletes the partial file when a chunk fails, then rethrows', async () => {
    const c = mkClient();
    c.postBinary
      .mockResolvedValueOnce({})
      .mockResolvedValueOnce({})
      .mockRejectedValueOnce(new Error('network died'));
    await expect(
      uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 3), '/a', 'big.bin', true),
    ).rejects.toThrowError('network died');
    expect(c.postJson).toHaveBeenCalledWith(expect.stringContaining('recycle'));
  });

  it('does not mask the original error if cleanup also fails', async () => {
    const c = mkClient();
    c.postBinary
      .mockResolvedValueOnce({})
      .mockResolvedValueOnce({})
      .mockRejectedValueOnce(new Error('network died'));
    c.postJson.mockRejectedValue(new Error('cleanup failed too'));
    await expect(
      uploadFile(c as never, Buffer.alloc(CHUNK_SIZE_BYTES * 3), '/a', 'big.bin', true),
    ).rejects.toThrowError('network died');
  });

  it('does not attempt cleanup when the small-file path fails', async () => {
    const c = mkClient();
    c.postBinary.mockRejectedValueOnce(new Error('denied'));
    await expect(
      uploadFile(c as never, Buffer.alloc(16), '/a', 'f.bin', true),
    ).rejects.toThrowError('denied');
    expect(c.postJson).not.toHaveBeenCalled();
  });
});
