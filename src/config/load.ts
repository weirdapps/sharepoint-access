// src/config/load.ts
//
// Configuration resolution. Precedence: CLI flag > env var > default.
//
// `host` deliberately has NO default: it identifies which tenant is
// addressed, and a default would silently point at the wrong one. See the
// recorded exception in CLAUDE.md.

import * as os from 'node:os';
import * as path from 'node:path';

import { CliError } from './errors';

export interface CliConfig {
  host: string;
  httpTimeoutMs: number;
  loginTimeoutMs: number;
  renewTimeoutMs: number;
  chromeChannel: string;
  sessionPath: string;
  profileDir: string;
  lockPath: string;
}

export interface ConfigOverrides {
  host?: string;
  httpTimeoutMs?: number;
  loginTimeoutMs?: number;
  renewTimeoutMs?: number;
  chromeChannel?: string;
  sessionPath?: string;
}

/** Defaults for the four plumbing settings. See the exception in CLAUDE.md. */
const DEFAULTS = {
  httpTimeoutMs: 30_000,
  loginTimeoutMs: 300_000,
  renewTimeoutMs: 30_000,
  chromeChannel: 'chrome',
} as const;

/**
 * Only a bare `<label>.sharepoint.com` hostname is accepted. Anchored at both
 * ends so `x.sharepoint.com.evil.com`, `evil.com@x.sharepoint.com`,
 * `x.sharepoint.com:8443` and `x.sharepoint.com/../y` are all rejected: this
 * value is interpolated straight into URLs that carry live session cookies.
 */
const HOST_RE = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?\.sharepoint\.com$/i;

export function stateDir(): string {
  return path.join(os.homedir(), '.sharepoint-cli');
}

export function defaultSessionPath(): string {
  return path.join(stateDir(), 'session.json');
}

function resolveInt(
  flag: number | undefined,
  envRaw: string | undefined,
  envName: string,
  fallback: number,
): number {
  if (flag !== undefined) return flag;
  if (envRaw === undefined) return fallback;
  const n = Number.parseInt(envRaw, 10);
  if (!Number.isFinite(n) || n <= 0) {
    throw new CliError('CONFIG_INVALID', `${envName} must be a positive integer, got "${envRaw}"`);
  }
  return n;
}

export function loadConfig(
  overrides: ConfigOverrides = {},
  env: NodeJS.ProcessEnv = process.env,
): CliConfig {
  const host = overrides.host ?? env.SHAREPOINT_CLI_HOST;
  if (!host) {
    throw new CliError(
      'CONFIG_MISSING',
      'host is required: pass --host <tenant>.sharepoint.com or set SHAREPOINT_CLI_HOST. It has no default because it identifies which tenant is addressed.',
    );
  }
  if (!HOST_RE.test(host)) {
    throw new CliError(
      'CONFIG_INVALID',
      `host must be a single <label>.sharepoint.com hostname, got "${host}"`,
    );
  }

  const sessionPath =
    overrides.sessionPath ?? env.SHAREPOINT_CLI_SESSION_FILE ?? defaultSessionPath();

  return {
    host: host.toLowerCase(),
    httpTimeoutMs: resolveInt(
      overrides.httpTimeoutMs,
      env.SHAREPOINT_CLI_HTTP_TIMEOUT_MS,
      'SHAREPOINT_CLI_HTTP_TIMEOUT_MS',
      DEFAULTS.httpTimeoutMs,
    ),
    loginTimeoutMs: resolveInt(
      overrides.loginTimeoutMs,
      env.SHAREPOINT_CLI_LOGIN_TIMEOUT_MS,
      'SHAREPOINT_CLI_LOGIN_TIMEOUT_MS',
      DEFAULTS.loginTimeoutMs,
    ),
    renewTimeoutMs: resolveInt(
      overrides.renewTimeoutMs,
      env.SHAREPOINT_CLI_RENEW_TIMEOUT_MS,
      'SHAREPOINT_CLI_RENEW_TIMEOUT_MS',
      DEFAULTS.renewTimeoutMs,
    ),
    chromeChannel:
      overrides.chromeChannel ?? env.SHAREPOINT_CLI_CHROME_CHANNEL ?? DEFAULTS.chromeChannel,
    sessionPath,
    profileDir: path.join(stateDir(), 'playwright-profile'),
    lockPath: path.join(stateDir(), '.browser.lock'),
  };
}
