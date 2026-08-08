#!/usr/bin/env node
// src/cli.ts
//
// Entry point for sharepoint-cli.
//
// SCAFFOLD ONLY. Commands are registered per the phasing in
// docs/design/project-design.md §11 and are not implemented yet:
//
//   Phase 1  login, auth-renew, auth-check, health-check
//   Phase 2  ls, get, search, libraries
//   Phase 3  mkdir, put
//
// Until then the CLI exposes --version and --help and exits with
// ExitCode.NotImplemented for anything else, so nothing silently
// pretends to work.

import { Command } from 'commander';

import { ExitCode } from './util/exit-codes';

const VERSION = '0.1.0';

/** Commands the design defines, with the phase that delivers each. */
const PLANNED: ReadonlyArray<{ name: string; phase: number; summary: string }> = [
  { name: 'login', phase: 1, summary: 'Interactive sign-in, capture and persist the session' },
  { name: 'auth-renew', phase: 1, summary: 'Headless silent re-capture via ESTSAUTHPERSISTENT' },
  { name: 'auth-check', phase: 1, summary: 'Probe read, write and search surfaces' },
  { name: 'health-check', phase: 1, summary: 'Same probes with per-probe timings, for cron' },
  { name: 'ls', phase: 2, summary: 'List a folder' },
  { name: 'get', phase: 2, summary: 'Download a file by path or URL' },
  { name: 'search', phase: 2, summary: 'Search files and sites' },
  { name: 'libraries', phase: 2, summary: 'List document libraries in a site' },
  { name: 'mkdir', phase: 3, summary: 'Create a folder' },
  { name: 'put', phase: 3, summary: 'Upload a file, auto-chunked above 10 MB' },
];

function notImplemented(name: string, phase: number): never {
  process.stderr.write(
    JSON.stringify({
      error: 'not_implemented',
      command: name,
      message: `"${name}" lands in phase ${phase}. See docs/design/project-design.md §11.`,
    }) + '\n',
  );
  process.exit(ExitCode.NotImplemented);
}

export function buildProgram(): Command {
  const program = new Command();

  program
    .name('sharepoint-cli')
    .description('CLI for SharePoint Online via a captured browser session')
    .version(VERSION);

  for (const { name, phase, summary } of PLANNED) {
    program
      .command(name)
      .description(`[phase ${phase}, not implemented] ${summary}`)
      // Swallow whatever the caller passes. Until these are built, arguments
      // must not turn into a commander usage error (exit 1): the caller needs
      // to see not_implemented (exit 7), not a misleading argument complaint.
      .allowUnknownOption(true)
      .allowExcessArguments(true)
      .action(() => notImplemented(name, phase));
  }

  return program;
}

if (require.main === module) {
  buildProgram().parse(process.argv);
}
