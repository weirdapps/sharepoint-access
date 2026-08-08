// src/commands/auth-check.ts
//
// Probes EVERY surface the CLI depends on, not just one.
//
// This is the direct lesson from teams-cli, whose auth-check probes Graph /me
// alone and therefore reports ok on a session whose chatsvc scope is dead.
// health-check reuses this exact function, so the two can never disagree the
// way teams-cli's pair does.

import type { SharepointClient } from '../http/client';

export type ProbeName = 'read' | 'write' | 'search';

export interface Probe {
  name: ProbeName;
  ok: boolean;
  detail: string;
  durationMs: number;
}

export interface AuthCheckResult {
  overall: 'ok' | 'degraded' | 'broken';
  probes: Probe[];
}

type Prober = Pick<SharepointClient, 'getJson' | 'contextInfo'>;

async function timed(name: ProbeName, fn: () => Promise<unknown>): Promise<Probe> {
  const started = Date.now();
  try {
    await fn();
    return { name, ok: true, detail: 'ok', durationMs: Date.now() - started };
  } catch (err) {
    return {
      name,
      ok: false,
      detail: (err as Error).message,
      durationMs: Date.now() - started,
    };
  }
}

export async function runAuthCheck(client: Prober, site = ''): Promise<AuthCheckResult> {
  // Probe the web the caller actually cares about: digests and folder access
  // are web-scoped, so a green root-web probe says nothing about a sub-site.
  const probes = await Promise.all([
    timed('read', () => client.getJson(`${site}/_api/web?$select=Title`)),
    // contextinfo is the write dependency: no digest means no writes, even
    // though every read still works.
    timed('write', () => client.contextInfo(site)),
    timed('search', () => client.getJson(`${site}/_api/search/query?querytext='test'&rowlimit=1`)),
  ]);

  const read = probes.find((p) => p.name === 'read');
  // Without reads nothing works, so a read failure is broken rather than
  // degraded. Anything else failing still leaves a usable subset.
  const overall = !read?.ok ? 'broken' : probes.every((p) => p.ok) ? 'ok' : 'degraded';

  return { overall, probes };
}
