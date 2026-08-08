// src/commands/login.ts
//
// Interactive first sign-in, and its headless sibling. Both take the browser
// lock so two Playwright instances can never race one profile directory.

import type { CliConfig } from '../config/load';
import { acquireLock } from '../auth/lock';
import { captureSession } from '../auth/capture';
import { saveSession } from '../session/store';

export interface LoginResult {
  status: 'ok';
  host: string;
  sessionFile: string;
  tokenExpiresAt: string;
  /** True when a Bearer was seen. Cookie-auth tenants report false. */
  bearerCaptured: boolean;
  durationMs: number;
}

async function capture(config: CliConfig, headless: boolean): Promise<LoginResult> {
  const started = Date.now();
  const release = await acquireLock(config.lockPath);
  try {
    const session = await captureSession({
      host: config.host,
      profileDir: config.profileDir,
      chromeChannel: config.chromeChannel,
      timeoutMs: headless ? config.renewTimeoutMs : config.loginTimeoutMs,
      headless,
    });
    await saveSession(config.sessionPath, session);
    return {
      status: 'ok',
      host: session.host,
      sessionFile: config.sessionPath,
      tokenExpiresAt: session.tokenExpiresAt,
      bearerCaptured: Boolean(session.bearer),
      durationMs: Date.now() - started,
    };
  } finally {
    await release();
  }
}

export function runLogin(config: CliConfig): Promise<LoginResult> {
  return capture(config, false);
}

export function runAuthRenew(config: CliConfig): Promise<LoginResult> {
  return capture(config, true);
}
