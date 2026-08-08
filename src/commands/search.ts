// src/commands/search.ts

import type { SharepointClient } from '../http/client';
import type { SpSearchResponse } from '../http/types';
import { CliError } from '../config/errors';
import { odataLiteral } from '../sharepoint/paths';

export interface SearchResult {
  query: string;
  totalRows?: number;
  results: Array<Record<string, string>>;
}

type Reader = Pick<SharepointClient, 'getJson'>;

const SELECT_PROPERTIES = 'Title,Path,FileType,LastModifiedTime,Size';
const MAX_ROWS = 500;

/** Property names that must never be taken from server-supplied cell keys. */
const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

export async function runSearch(
  client: Reader,
  query: string,
  rows: number,
): Promise<SearchResult> {
  if (typeof query !== 'string' || query.trim().length === 0) {
    throw new CliError('CONFIG_INVALID', 'search query must be a non-empty string');
  }
  if (!Number.isInteger(rows) || rows <= 0 || rows > MAX_ROWS) {
    throw new CliError('CONFIG_INVALID', `rows must be an integer between 1 and ${MAX_ROWS}`);
  }

  // Same two-layer escaping as paths: the query sits inside an OData string
  // literal, so an apostrophe would otherwise terminate it early.
  const q = encodeURIComponent(odataLiteral(query));
  const url =
    `/_api/search/query?querytext='${q}'&rowlimit=${rows}` +
    `&selectproperties='${encodeURIComponent(SELECT_PROPERTIES)}'`;

  const body = await client.getJson<SpSearchResponse>(url);
  const relevant = body.PrimaryQueryResult?.RelevantResults;
  const results = (relevant?.Table?.Rows ?? []).map((row) => {
    // Null-prototype: Cell.Key is server-supplied and ends up as a property
    // name, so a row keyed "__proto__" or "constructor" would otherwise reach
    // Object.prototype. A managed-path document title can influence indexed
    // properties, so this is reachable input, not just theory.
    const out = Object.create(null) as Record<string, string>;
    for (const cell of row.Cells ?? []) {
      if (typeof cell.Key === 'string' && typeof cell.Value === 'string') {
        if (FORBIDDEN_KEYS.has(cell.Key)) continue;
        out[cell.Key] = cell.Value;
      }
    }
    // Back to an ordinary object so JSON.stringify and deep-equality behave.
    return { ...out };
  });

  return {
    query,
    ...(typeof relevant?.TotalRows === 'number' ? { totalRows: relevant.TotalRows } : {}),
    results,
  };
}
