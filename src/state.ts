// Durable state for Mastra.
//
// Mastra keeps a suspended run's snapshot in its storage adapter. With the in-memory store that
// snapshot dies with the process, which made "draft now, let the human mark the sheet after lunch"
// impossible — the exact thing this project is about. `LibSQLStore` over a local file fixes it.
//
// This is NOT `src/db.ts`. That file is our own ledger (dedupe + audit). This one holds Mastra's
// internal run state, and we never query it directly.

import { Mastra } from '@mastra/core';
import { LibSQLStore } from '@mastra/libsql';
import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

export const DEFAULT_MASTRA_DB = '.data/mastra.db';

/** A Mastra instance whose workflow snapshots survive a restart. */
export function durableMastra(file = process.env.SHEETFLOW_MASTRA_DB ?? DEFAULT_MASTRA_DB): Mastra {
  const path = resolve(file);
  mkdirSync(dirname(path), { recursive: true });
  return new Mastra({
    storage: new LibSQLStore({ id: 'sheetflow', url: `file:${path}` }),
  });
}
