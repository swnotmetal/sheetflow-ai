// The pipeline's own memory — NOT a second database.
//
// Business data lives in the user's workbook. This SQLite file only stores:
//   1. the state of a suspended workflow (so a restart resumes, not replays), and
//   2. an append-only audit trail ("who approved what, when").
//
// The ledger is synchronous, which matches a single-file, single-writer CLI perfectly. It uses
// `node:sqlite` rather than better-sqlite3 on purpose: better-sqlite3 is a native module that has to
// be compiled, or matched to the exact Node ABI, on the machine that runs it — which a shop owner's
// PC cannot do. `node:sqlite` ships inside Node, so the whole tool packages into one executable.

import { DatabaseSync } from 'node:sqlite';
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
  // SQLite creates a file but not a directory, and `.data/` is gitignored — so on a fresh clone the
  // very first run failed with "the directory does not exist".
  if (file !== ':memory:') mkdirSync(dirname(file), { recursive: true });
  const db = new DatabaseSync(file);
  db.exec('PRAGMA journal_mode = WAL');
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
    -- Things the shop owner must see, but which cannot always be written when they happen: the most
    -- common failure is "the workbook is open in Excel", and that is exactly when the workbook
    -- cannot be written. So they queue here and the next pass that finds the file free projects
    -- them into the Log sheet. The sheet is the only interface, so a failure has to reach it.
    CREATE TABLE IF NOT EXISTS sheet_notices (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      run_id TEXT NOT NULL,
      message_id TEXT NOT NULL,
      when_label TEXT NOT NULL,
      mail_label TEXT NOT NULL,
      did TEXT NOT NULL,
      result TEXT NOT NULL,
      projected_at TEXT,
      -- One line per mail per failure: a file left open for an hour must not fill the sheet.
      UNIQUE (message_id, result)
    );
  `);
  // Which mail this run is, as a human reads it ("2026-10-30 09:12 UTC · ACME <orders@acme>"). The
  // Log sheet is written from it, including when a receipt is repaired long after the run finished,
  // when no suspended payload is left to ask. Added after the first schema, so it is migrated in.
  const columns = db.prepare(`PRAGMA table_info(workflow_runs)`).all() as { name: string }[];
  if (!columns.some((column) => column.name === 'mail_label')) {
    db.exec(`ALTER TABLE workflow_runs ADD COLUMN mail_label TEXT`);
  }
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
    // node:sqlite types named parameters as a plain record; the row is exactly those columns.
    .run(run as unknown as Record<string, string>);
  return Number(info.changes) > 0;
}

export function setRunStatus(db: Db, id: string, status: RunStatus): void {
  db.prepare(`UPDATE workflow_runs SET status = ?, updated_at = ? WHERE id = ?`).run(
    status,
    new Date().toISOString(),
    id,
  );
}

export function getRun(db: Db, messageId: string): RunRow | undefined {
  return db.prepare(`SELECT * FROM workflow_runs WHERE message_id = ?`).get(messageId) as unknown as
    | RunRow
    | undefined;
}

/** Runs a previous pass left waiting on a human. */
export function listRunsByStatus(db: Db, status: RunStatus): RunRow[] {
  return db
    .prepare(`SELECT * FROM workflow_runs WHERE status = ? ORDER BY created_at`)
    .all(status) as unknown as RunRow[];
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

/** Remember the mail as a human reads it, so the sheet never has to show a Message-ID. */
export function setMailLabel(db: Db, id: string, label: string): void {
  db.prepare(`UPDATE workflow_runs SET mail_label = ? WHERE id = ?`).run(label, id);
}

/** The mail as a human reads it, or `''` when the run predates this column. */
export function getMailLabel(db: Db, id: string): string {
  const row = db.prepare(`SELECT mail_label FROM workflow_runs WHERE id = ?`).get(id) as
    | { mail_label: string | null }
    | undefined;
  return row?.mail_label ?? '';
}

/** One line the shop owner has to see, queued until the workbook can be written. */
export interface SheetNotice {
  id: number;
  message_id: string;
  when_label: string;
  mail_label: string;
  did: string;
  result: string;
}

/**
 * Queue a line for the `Log` sheet. Idempotent per (mail, failure), because the usual cause — the
 * workbook being open in Excel — otherwise repeats on every pass for as long as it stays open.
 */
export function notice(
  db: Db,
  entry: { runId: string; messageId: string; when: string; mail: string; did: string; result: string },
): void {
  db.prepare(
    `INSERT OR IGNORE INTO sheet_notices (run_id, message_id, when_label, mail_label, did, result)
     VALUES (?, ?, ?, ?, ?, ?)`,
  ).run(entry.runId, entry.messageId, entry.when, entry.mail, entry.did, entry.result);
}

/** Notices the workbook has not been told about yet, oldest first. */
export function pendingNotices(db: Db): SheetNotice[] {
  return db
    .prepare(
      `SELECT id, message_id, when_label, mail_label, did, result
       FROM sheet_notices WHERE projected_at IS NULL ORDER BY id`,
    )
    .all() as unknown as SheetNotice[];
}

export function markNoticesProjected(db: Db, ids: number[]): void {
  if (ids.length === 0) return;
  const at = new Date().toISOString();
  const mark = db.prepare(`UPDATE sheet_notices SET projected_at = ? WHERE id = ?`);
  for (const id of ids) mark.run(at, id);
}

/**
 * Drop a notice that no longer describes reality — the problem it reported was repaired before the
 * sheet got to read it. Writing both would have the Log say "recorded ✓" and "not written yet" about
 * the same row.
 */
export function supersedeNotices(db: Db, runId: string, result: string): void {
  db.prepare(
    `UPDATE sheet_notices SET projected_at = ?
     WHERE run_id = ? AND result = ? AND projected_at IS NULL`,
  ).run(new Date().toISOString(), runId, result);
}
