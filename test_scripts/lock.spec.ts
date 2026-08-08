import { describe, expect, it } from 'vitest';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';

import { acquireLock } from '../src/auth/lock';

function tmpLock(): string {
  return path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'spl-')), 'nested', '.browser.lock');
}

describe('acquireLock', () => {
  it('creates the lock file 0600 and its parent 0700', async () => {
    const p = tmpLock();
    const release = await acquireLock(p);
    expect(fs.statSync(p).mode & 0o777).toBe(0o600);
    expect(fs.statSync(path.dirname(p)).mode & 0o777).toBe(0o700);
    await release();
  });

  it('writes the owning PID', async () => {
    const p = tmpLock();
    const release = await acquireLock(p);
    expect(fs.readFileSync(p, 'utf8').trim()).toBe(String(process.pid));
    await release();
  });

  it('refuses a second acquire while a live process holds it', async () => {
    const p = tmpLock();
    const release = await acquireLock(p);
    await expect(acquireLock(p)).rejects.toThrowError(/holds the lock/);
    await release();
  });

  it('names sharepoint-cli, not the repo it was ported from', async () => {
    const p = tmpLock();
    const release = await acquireLock(p);
    await expect(acquireLock(p)).rejects.toThrowError(/sharepoint-cli/);
    await release();
  });

  it('removes the file on release', async () => {
    const p = tmpLock();
    const release = await acquireLock(p);
    await release();
    expect(fs.existsSync(p)).toBe(false);
  });

  it('is idempotent on repeated release', async () => {
    const p = tmpLock();
    const release = await acquireLock(p);
    await release();
    await expect(release()).resolves.toBeUndefined();
  });

  it('reclaims a lock whose PID is dead', async () => {
    const p = tmpLock();
    fs.mkdirSync(path.dirname(p), { recursive: true });
    // PID 2^22 is above the macOS/Linux maximum, so it cannot be running.
    fs.writeFileSync(p, '4194304\n');
    const release = await acquireLock(p);
    expect(fs.readFileSync(p, 'utf8').trim()).toBe(String(process.pid));
    await release();
  });

  it('reclaims a lock whose contents are garbage', async () => {
    const p = tmpLock();
    fs.mkdirSync(path.dirname(p), { recursive: true });
    fs.writeFileSync(p, 'not-a-pid\n');
    const release = await acquireLock(p);
    await release();
    expect(fs.existsSync(p)).toBe(false);
  });

  it('reclaims an empty lock file', async () => {
    const p = tmpLock();
    fs.mkdirSync(path.dirname(p), { recursive: true });
    fs.writeFileSync(p, '');
    const release = await acquireLock(p);
    await release();
    expect(fs.existsSync(p)).toBe(false);
  });
});
