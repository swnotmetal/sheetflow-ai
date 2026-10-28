// The pipeline's own memory — NOT a second database.
//
// Business data lives in the user's workbook. This SQLite file only stores:
//   1. the state of a suspended workflow (so a restart resumes, not replays), and
//   2. an append-only audit trail ("who approved what, when").
//
// better-sqlite3 is synchronous, which matches a single-file, single-writer CLI perfectly.

import Database from 'better-sqlite3';
import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

export interface RunRow {
  id: string;
  thread_id: string;
  message_id: string;
  status: RunStatus;
  created_at: string;
  updated_at: string;
}

/**
 * `failed` exists because "the model never answered" is not a human decision and must not be
 * recorded as one. `unmatched` means the mail named nothing the sheet contains. `ignored` means
 * the mail was not an order at all. `sent` stays for the outbound transport that is not wired yet.
 */
export type RunStatus =
  | 'pending'
  | 'approved'
  | 'rejected'
  | 'sent'
  | 'failed'
  | 'unmatched'
  | 'ignored';

export function openDb(file = resolve(process.env.SHEETFLOW_DATA_DIR ?? '.data', 'sheetflow.sqlite')) {
  // better-sqlite3 creates a file but not a directory, and `.data/` is gitignored — so on a fresh
  // clone the very first run failed with "the directory does not exist".
  if (file !== ':memory:') mkdirSync(dirname(file), { recursive: true });
  const db = new Database(file);
  db.pragma('journal_mode = WAL');
  db.exec(`
    CREATE TABLE IF NOT EXISTS workflow_runs (
      id TEXT PRIMARY KEY,
      thread_id TEXT NOT NULL,
      message_id TEXT NOT NULL UNIQUE,
      status TEXT NOT NULL DEFAULT 'pending',
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS audit_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      run_id TEXT NOT NULL,
      actor TEXT NOT NULL,
      action TEXT NOT NULL,
      at TEXT NOT NULL
    );
  `);
  return db;
}

export type Db = ReturnType<typeof openDb>;

// UNIQUE(message_id) is the dedupe: pulling the same row twice cannot create two runs.
// Returns whether this call created the row — the caller's cue to do real work or skip.
export function upsertRun(db: Db, run: RunRow): boolean {
  const info = db
    .prepare(
      `INSERT INTO workflow_runs (id, thread_id, message_id, status, created_at, updated_at)
       VALUES (@id, @thread_id, @message_id, @status, @created_at, @updated_at)
       ON CONFLICT(message_id) DO NOTHING`,
    )
    .run(run);
  return info.changes > 0;
}

export function setRunStatus(db: Db, id: string, status: RunStatus): void {
  db.prepare(`UPDATE workflow_runs SET status = ?, updated_at = ? WHERE id = ?`).run(
    status,
    new Date().toISOString(),
    id,
  );
}

export function getRun(db: Db, messageId: string): RunRow | undefined {
  return db.prepare(`SELECT * FROM workflow_runs WHERE message_id = ?`).get(messageId) as
    | RunRow
    | undefined;
}

/** Runs a previous pass left waiting on a human. */
export function listRunsByStatus(db: Db, status: RunStatus): RunRow[] {
  return db.prepare(`SELECT * FROM workflow_runs WHERE status = ? ORDER BY created_at`).all(status) as RunRow[];
}

/**
 * Forget a run. Used when the pass could not do anything at all (the workbook was locked), so the
 * message stays eligible instead of being deduped away forever.
 */
export function deleteRun(db: Db, id: string): void {
  db.prepare(`DELETE FROM workflow_runs WHERE id = ?`).run(id);
}

export function audit(db: Db, runId: string, actor: string, action: string): void {
  db.prepare(`INSERT INTO audit_log (run_id, actor, action, at) VALUES (?, ?, ?, ?)`).run(
    runId,
    actor,
    action,
    new Date().toISOString(),
  );
}
