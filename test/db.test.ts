// The ledger, including the first-run case that broke for the owner: `.data/` is gitignored, so on a
// fresh clone the directory does not exist and better-sqlite3 refuses to create the file.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { getRun, listRunsByStatus, openDb, setRunStatus, upsertRun } from '../src/db.js';

test('a database in a directory that does not exist yet is created, not rejected', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-db-'));
  const file = join(dir, 'nested', 'deeper', 'sheetflow.sqlite');
  try {
    const db = openDb(file);
    assert.equal(existsSync(file), true);
    db.close();
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('one message is one run — the second insert is refused', () => {
  const db = openDb(':memory:');
  const row = {
    id: 'run-1',
    thread_id: '<m@x>',
    message_id: '<m@x>',
    status: 'pending' as const,
    created_at: '2026-10-28T09:00:00.000Z',
    updated_at: '2026-10-28T09:00:00.000Z',
  };

  assert.equal(upsertRun(db, row), true);
  assert.equal(upsertRun(db, { ...row, id: 'run-2' }), false, 'UNIQUE(message_id) holds');

  setRunStatus(db, 'run-1', 'approved');
  assert.equal(getRun(db, '<m@x>')?.status, 'approved');
  assert.deepEqual(listRunsByStatus(db, 'pending'), []);
  assert.deepEqual(listRunsByStatus(db, 'approved').map((run) => run.id), ['run-1']);
  db.close();
});
