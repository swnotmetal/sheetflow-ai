// The pipeline's own memory — NOT a second database.
//
// Business data lives in the user's workbook. This SQLite file only stores:
//   1. the state of a suspended workflow (so a restart resumes, not replays), and
//   2. an append-only audit trail ("who approved what, when").
//
// better-sqlite3 is synchronous, which matches a single-file, single-writer CLI perfectly.

import Database from 'better-sqlite3';
import { resolve } from 'node:path';

export interface RunRow {
  id: string;
  thread_id: string;
  message_id: string;
  status: 'pending' | 'approved' | 'rejected' | 'sent';
  created_at: string;
  updated_at: string;
}

export function openDb(file = resolve(process.env.SHEETFLOW_DATA_DIR ?? '.data', 'sheetflow.sqlite')) {
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
export function upsertRun(db: Db, run: RunRow): void {
  db.prepare(
    `INSERT INTO workflow_runs (id, thread_id, message_id, status, created_at, updated_at)
     VALUES (@id, @thread_id, @message_id, @status, @created_at, @updated_at)
     ON CONFLICT(message_id) DO NOTHING`,
  ).run(run);
}

export function audit(db: Db, runId: string, actor: string, action: string): void {
  db.prepare(`INSERT INTO audit_log (run_id, actor, action, at) VALUES (?, ?, ?, ?)`).run(
    runId,
    actor,
    action,
    new Date().toISOString(),
  );
}
