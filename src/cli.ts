#!/usr/bin/env node
// src/cli.ts
//
// Commander wiring only. Each command is a pure function elsewhere returning a
// plain object; this layer resolves config, builds the client, and formats.
// Errors become a JSON payload on stderr plus a mapped exit code.

import { Command, Option } from 'commander';

import { loadConfig, type CliConfig, type ConfigOverrides } from './config/load';
import { CliError } from './config/errors';
import { SharepointClient } from './http/client';
import { DigestCache } from './sharepoint/digest';
import { loadSession } from './session/store';
import { emit, emitError, toExit } from './output/json';
import { ExitCode } from './util/exit-codes';

import { runLogin, runAuthRenew } from './commands/login';
import { runAuthCheck } from './commands/auth-check';
import { runLs } from './commands/ls';
import { runGet } from './commands/get';
import { runLibraries } from './commands/libraries';
import { runSearch } from './commands/search';
import { runMkdir } from './commands/mkdir';
import { runPut } from './commands/put';

const VERSION = '0.1.0';

interface GlobalOpts {
  host?: string;
  site?: string;
  timeout?: string;
  loginTimeout?: string;
  sessionFile?: string;
  chromeChannel?: string;
}

export function parseIntOption(raw: string, flag: string): number {
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n <= 0 || String(n) !== raw.trim()) {
    throw new CliError('CONFIG_INVALID', `${flag} must be a positive integer, got "${raw}"`);
  }
  return n;
}

function siteFrom(program: Command): string | undefined {
  return program.opts<GlobalOpts>().site;
}

function configFrom(program: Command): CliConfig {
  const g = program.opts<GlobalOpts>();
  const overrides: ConfigOverrides = {};
  if (g.host) overrides.host = g.host;
  if (g.sessionFile) overrides.sessionPath = g.sessionFile;
  if (g.chromeChannel) overrides.chromeChannel = g.chromeChannel;
  if (g.timeout) overrides.httpTimeoutMs = parseIntOption(g.timeout, '--timeout');
  if (g.loginTimeout) overrides.loginTimeoutMs = parseIntOption(g.loginTimeout, '--login-timeout');
  return loadConfig(overrides);
}

async function clientFrom(config: CliConfig): Promise<SharepointClient> {
  const session = await loadSession(config.sessionPath);
  if (!session) {
    throw new CliError(
      'AUTH_REQUIRED',
      `no session at ${config.sessionPath}: run "sharepoint-cli login --host ${config.host}"`,
    );
  }
  // The stored session may have been captured against the other host in the
  // pair. Cookies cover both (parent-domain scope), so retarget the host
  // rather than forcing a pointless re-login.
  const client = new SharepointClient(
    { ...session, host: config.host },
    { httpTimeoutMs: config.httpTimeoutMs },
  );
  const digest = new DigestCache(client);
  client.setDigestProvider((web, force) => digest.get(web, force));
  return client;
}

/** Single error boundary: every action body runs inside this. */
async function run(fn: () => Promise<unknown>): Promise<void> {
  try {
    emit(await fn());
    process.exitCode = ExitCode.Success;
  } catch (err) {
    const { code, payload } = toExit(err);
    emitError(payload);
    process.exitCode = code;
  }
}

export function buildProgram(): Command {
  const program = new Command();

  program
    .name('sharepoint-cli')
    .description(
      'CLI for SharePoint Online and OneDrive for Business via a captured browser session',
    )
    .version(VERSION)
    .addOption(
      new Option(
        '--host <host>',
        'tenant host, e.g. <tenant>.sharepoint.com or <tenant>-my.sharepoint.com',
      ).env('SHAREPOINT_CLI_HOST'),
    )
    .option('--timeout <ms>', 'per-request HTTP timeout')
    .option('--login-timeout <ms>', 'max wait for interactive sign-in')
    .option('--session-file <path>', 'override the session file path')
    .option(
      '--site <path>',
      'explicit web/sub-site path, e.g. /sites/foo. Derived from the target path when omitted; needed only for webs nested deeper than two segments',
    )
    .option('--chrome-channel <name>', 'Playwright Chrome channel');

  program
    .command('login')
    .description('Interactive sign-in; captures and persists the session')
    .action(() => run(() => runLogin(configFrom(program))));

  program
    .command('auth-renew')
    .description('Headless silent re-capture of the session')
    .action(() => run(() => runAuthRenew(configFrom(program))));

  program
    .command('auth-check')
    .description('Probe the read, write and search surfaces')
    .action(() =>
      run(async () => runAuthCheck(await clientFrom(configFrom(program)), siteFrom(program) ?? '')),
    );

  program
    .command('health-check')
    .description('Same probes as auth-check, with timings, for cron')
    .action(() =>
      run(async () => runAuthCheck(await clientFrom(configFrom(program)), siteFrom(program) ?? '')),
    );

  program
    .command('ls')
    .description('List a folder')
    .argument('<path>', 'server-relative path')
    .action((path: string) =>
      run(async () => runLs(await clientFrom(configFrom(program)), path, siteFrom(program))),
    );

  program
    .command('get')
    .description('Download a file by server-relative path or absolute SharePoint URL')
    .argument('<path-or-url>')
    .option('--out <file>', 'write the bytes to this file')
    .action((pathOrUrl: string, opts: { out?: string }) =>
      run(async () =>
        runGet(await clientFrom(configFrom(program)), pathOrUrl, opts.out, siteFrom(program)),
      ),
    );

  program
    .command('libraries')
    .description('List document libraries in the site')
    .action(() =>
      run(async () => runLibraries(await clientFrom(configFrom(program)), siteFrom(program))),
    );

  program
    .command('search')
    .description('Search files and sites')
    .argument('<query>')
    .option('--rows <n>', 'max results', '20')
    .action((query: string, opts: { rows: string }) =>
      run(async () =>
        runSearch(
          await clientFrom(configFrom(program)),
          query,
          parseIntOption(opts.rows, '--rows'),
        ),
      ),
    );

  program
    .command('mkdir')
    .description('Create one folder; parents must already exist')
    .argument('<path>')
    .action((path: string) =>
      run(async () => runMkdir(await clientFrom(configFrom(program)), path, siteFrom(program))),
    );

  program
    .command('put')
    .description('Upload a file, auto-chunked above 10 MB')
    .argument('<local>')
    .argument('<remote-folder>')
    .option('--overwrite', 'replace an existing file of the same name', false)
    .action((local: string, remoteFolder: string, opts: { overwrite: boolean }) =>
      run(async () =>
        runPut(
          await clientFrom(configFrom(program)),
          local,
          remoteFolder,
          opts.overwrite,
          siteFrom(program),
        ),
      ),
    );

  return program;
}

if (require.main === module) {
  buildProgram()
    .parseAsync(process.argv)
    .catch((err) => {
      const { code, payload } = toExit(err);
      emitError(payload);
      process.exit(code);
    });
}
