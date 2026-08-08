import { describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { parseSession, serializeSession, loadSession, saveSession } from '../src/session/store';

const valid = {
  version: 1 as const,
  host: 'x.sharepoint.com',
  cookies: 'FedAuth=aaa; rtFa=bbb',
  capturedAt: '2026-08-08T00:00:00.000Z',
  tokenExpiresAt: '2026-08-13T00:00:00.000Z',
};

describe('parseSession', () => {
  it('accepts a cookie-only session with no bearer', () => {
    expect(parseSession(JSON.stringify(valid)).bearer).toBeUndefined();
  });

  it('accepts a session with a bearer', () => {
    expect(parseSession(JSON.stringify({ ...valid, bearer: 'jwt' })).bearer).toBe('jwt');
  });

  it('rejects an unsupported version', () => {
    expect(() => parseSession(JSON.stringify({ ...valid, version: 2 }))).toThrowError(/version/);
  });

  it('rejects a missing cookies field', () => {
    const { cookies: _drop, ...rest } = valid;
    expect(() => parseSession(JSON.stringify(rest))).toThrowError(/cookies/);
  });

  it('rejects an empty cookies field', () => {
    expect(() => parseSession(JSON.stringify({ ...valid, cookies: '' }))).toThrowError(/cookies/);
  });

  it('rejects a non-string bearer', () => {
    expect(() => parseSession(JSON.stringify({ ...valid, bearer: 42 }))).toThrowError(/bearer/);
  });

  it('rejects malformed JSON', () => {
    expect(() => parseSession('{oops')).toThrowError(/JSON/);
  });

  it('rejects a JSON array', () => {
    expect(() => parseSession('[]')).toThrowError();
  });

  it('rejects JSON null', () => {
    expect(() => parseSession('null')).toThrowError(/object/);
  });
});

describe('saveSession / loadSession', () => {
  it('round-trips and writes the file 0600 in a 0700 directory', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    const p = path.join(dir, 'nested', 'session.json');
    await saveSession(p, valid);
    expect(await loadSession(p)).toEqual(valid);
    expect(fs.statSync(p).mode & 0o777).toBe(0o600);
    expect(fs.statSync(path.dirname(p)).mode & 0o777).toBe(0o700);
  });

  it('returns null when the file does not exist', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    expect(await loadSession(path.join(dir, 'absent.json'))).toBeNull();
  });

  it('does not leave a .tmp file behind', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    const p = path.join(dir, 'session.json');
    await saveSession(p, valid);
    expect(fs.existsSync(p + '.tmp')).toBe(false);
  });

  it('overwrites an existing session without widening its mode', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    const p = path.join(dir, 'session.json');
    await saveSession(p, valid);
    await saveSession(p, { ...valid, cookies: 'FedAuth=new' });
    expect((await loadSession(p))?.cookies).toBe('FedAuth=new');
    expect(fs.statSync(p).mode & 0o777).toBe(0o600);
  });

  it('propagates a parse error for a corrupt file rather than returning null', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spc-'));
    const p = path.join(dir, 'session.json');
    fs.writeFileSync(p, '{corrupt');
    await expect(loadSession(p)).rejects.toThrowError(/JSON/);
  });
});

describe('serializeSession', () => {
  it('produces parseable output', () => {
    expect(parseSession(serializeSession(valid))).toEqual(valid);
  });
});
