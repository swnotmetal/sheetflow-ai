// Durability, in its own process on purpose.
//
// The test resumes a run that a *different* Mastra instance suspended, over the same on-disk
// LibSQL store — the property the whole project depends on. It lives alone because mixing the
// libsql worker and better-sqlite3 handles from the other test file makes Node abort during
// environment teardown (`RemoveEnvironmentCleanupHook`); the CLI hit the same thing and fixes it by
// closing both handles before exit.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Mastra } from '@mastra/core';
import { LibSQLStore } from '@mastra/libsql';
import { buildWorkflow } from '../src/workflow.js';
import { resolveProfile, type Extraction } from '../src/model.js';
import { getRun, openDb } from '../src/db.js';
import { pollInbox } from '../src/inbox.js';
import type { Mail, Mailbox } from '../src/mail.js';
import type { Workbook } from '../src/excel.js';

const MAIL: Mail = {
  messageId: '<durable-1@acer.example>',
  from: 'orders@acer.example',
  subject: 'PO 4500 — restock request',
  body: 'Please ship 30 units of steel bracket to our depot.',
  receivedAt: '2026-10-28T08:00:00.000Z',
};

const EXTRACTION: Extraction = {
  intent: 'order',
  po: ['4500'],
  lines: [{ description: 'steel bracket', quantity: '30' }],
};

test('a run suspended by one process is resumed by a fresh one on the same file store', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-state-'));
  const url = `file:${join(dir, 'mastra.db')}`;
  const db = openDb(join(dir, 'sheetflow.sqlite'));

  const drafts = new Map<number, string>();
  const marks = new Map<number, 'approved' | 'rejected'>();
  const book: Workbook = {
    async ensureInterface() {},
    async products() {
      return [{ sku: 'steel bracket', name: 'steel bracket' }];
    },
    async resolveRows(lines) {
      const rows = new Map<number, number>();
      const missing: string[] = [];
      lines.forEach((line, at) => {
        if (line.description === 'steel bracket') rows.set(at, 4);
        else missing.push(line.description);
      });
      return { rows, missing };
    },
    async writeDraft(value, cell) {
      drafts.set(cell.row, value);
    },
    async readApproval(row) {
      return marks.get(row) ?? 'pending';
    },
    async clearDraft(row) {
      drafts.delete(row);
    },
    async markResults() {},
    async readResult() {
      return '';
    },
    async appendLog() {},
    async revertLastWrite() {
      drafts.clear();
    },
  };

  const instances: Mastra[] = [];
  // Two separate instances over the same file = two processes, two lifetimes.
  const makeWorkflow = () => {
    const mastra = new Mastra({ storage: new LibSQLStore({ id: 'sheetflow-test', url }) });
    instances.push(mastra);
    return buildWorkflow({
      mastra,
      book,
      extract: async () => EXTRACTION,
      provider: resolveProfile('openai'),
    });
  };
  const mailbox: Mailbox = { async fetch() { return [MAIL]; } };
  const pass = (workflow: ReturnType<typeof makeWorkflow>) =>
    pollInbox({ db, book, workflow, mailbox, pollMs: 1, timeoutMs: 5 });

  try {
    const first = await pass(makeWorkflow());
    assert.equal(first[0].action, 'pending', 'no mark yet: left suspended on disk');

    marks.set(4, 'approved');
    const second = await pass(makeWorkflow());

    assert.equal(second[0].action, 'approved');
    assert.equal(getRun(db, MAIL.messageId)?.status, 'approved');
  } finally {
    // Both native handles must go before Node tears the environment down.
    for (const mastra of instances) await mastra.shutdown({ drainTimeout: 2_000 });
    db.close();
  }
});
