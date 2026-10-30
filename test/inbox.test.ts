// The driver: dedupe, waiting across processes, and the audit trail.
//
// Offline and deterministic: a fake mailbox, in-memory or file-backed SQLite, a mock model, and a
// sheet whose marks appear between passes — which is exactly the real sequence.

import assert from 'node:assert/strict';
import test, { after } from 'node:test';
import { Mastra } from '@mastra/core';
import { InMemoryStore } from '@mastra/core/storage';
import { buildWorkflow } from '../src/workflow.js';
import { resolveProfile, type Extraction } from '../src/model.js';
import { openDb, getRun, pendingNotices, type Db } from '../src/db.js';
import { pollInbox, forget, forgetWaiting } from '../src/inbox.js';
import type { Mail, Mailbox } from '../src/mail.js';
import type { Workbook } from '../src/excel.js';

const MAIL: Mail = {
  messageId: '<inbox-1@acer.example>',
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

/** A mailbox that keeps returning the same message, like a read-only IMAP poll would. */
function fakeMailbox(mails: Mail[] = [MAIL]): Mailbox {
  return {
    async fetch() {
      return [...mails];
    },
  };
}

test('a run left undecided by an earlier pass does not hold up new mail', async () => {
  const db = openDb(':memory:');
  const { book, drafts, marks } = fakeBook({ 'steel bracket': 4, 'brass fitting': 6 });
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async (_target: unknown, text: string) =>
      /brass/.test(text)
        ? { intent: 'order' as const, lines: [{ description: 'brass fitting', quantity: '12' }] }
        : EXTRACTION,
    provider: resolveProfile('openai'),
  });
  const pass = (mails: Mail[], timeoutMs: number) =>
    pollInbox({ db, book, workflow, mailbox: fakeMailbox(mails), pollMs: 1, timeoutMs });

  // First pass: nobody answers row 4, so the run stays suspended in `pending`.
  const first = await pass([MAIL], 5);
  assert.equal(first[0].action, 'pending');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'pending');

  // Second pass has a ten-minute budget. The old run must not spend it: the new mail is handled now.
  const second: Mail = {
    ...MAIL,
    messageId: '<inbox-2@acer.example>',
    body: 'Please send 12 pcs of brass fitting.',
  };
  marks.set(6, 'approved');
  const startedAt = Date.now();
  const outcomes = await pass([second], 600_000);

  assert.ok(Date.now() - startedAt < 5_000, 'the pass did not sit on the old run');
  assert.equal(drafts.get(6), '12 · brass fitting');
  assert.equal(outcomes.find((outcome) => outcome.messageId === second.messageId)?.action, 'approved');
  assert.equal(outcomes.find((outcome) => outcome.messageId === MAIL.messageId)?.action, 'pending');
});

test('forgetting a waiting run lets the same mail be drafted again', async () => {
  const { db, pass } = harness();

  const first = await pass();
  assert.equal(first[0].action, 'pending');

  const forgotten = await forgetWaiting(db);
  assert.deepEqual(forgotten, [MAIL.messageId]);
  assert.equal(getRun(db, MAIL.messageId), undefined);

  const second = await pass();
  assert.equal(second[0].action, 'pending', 'drafted again instead of skipped');
  const audited = db
    .prepare(`SELECT action FROM audit_log`)
    .all()
    .map((row) => (row as { action: string }).action);
  assert.ok(audited.some((action) => action.startsWith('abandoned')), 'the ledger says it was discarded');
});

test('a rejected row is still recorded when Excel refuses the cleanup write', async () => {
  const db = openDb(':memory:');
  // This is the exact failure from 2026-10-28: `clearDraft` threw EPERM, the run was marked failed
  // and the human's `no` was lost. Clearing a rejected draft is now the driver's write-back, so a
  // locked file costs the visual cleanup and nothing else.
  const { book } = fakeBook(undefined, { lockedClears: true });
  const decisions = new Map<number, 'approved' | 'rejected'>([[4, 'rejected']]);
  const answering: Workbook = { ...book, readApproval: async (row) => decisions.get(row) ?? 'pending' };
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book: answering,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({
    db,
    book: answering,
    workflow,
    mailbox: fakeMailbox(),
    pollMs: 1,
    timeoutMs: 5_000,
    writeBack: { retries: 0, pauseMs: 1 },
  });

  assert.equal(outcomes[0].action, 'rejected', 'the decision survives the locked cleanup');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'rejected');
});

test('a run that failed can be forgotten so the mail is drafted again', async () => {
  const { db, pass } = harness();
  await pass();
  const runId = getRun(db, MAIL.messageId)!.id;
  db.prepare(`UPDATE workflow_runs SET status = 'failed' WHERE id = ?`).run(runId);

  const forgotten = await forget(db, ['failed'], 'test');
  assert.deepEqual(forgotten, [MAIL.messageId]);
  assert.equal(getRun(db, MAIL.messageId), undefined);

  const again = await pass();
  assert.equal(again[0].action, 'pending', 'the mail was processed again');
});

test('a receipt the workbook never received is written on the next pass', async () => {
  const db = openDb(':memory:');
  const { book, marks, results, log } = fakeBook(undefined, { lockedWrites: true });
  marks.set(4, 'approved');
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  // First pass: the human approved, but Excel held the file so nothing reached the sheet.
  await pollInbox({
    db,
    book,
    workflow,
    mailbox: fakeMailbox(),
    pollMs: 1,
    timeoutMs: 5_000,
    writeBack: { retries: 0, pauseMs: 1 },
  });
  assert.equal(getRun(db, MAIL.messageId)?.status, 'approved', 'the decision is safe in the ledger');
  assert.equal(results.size, 0, 'the sheet saw nothing yet');

  // Second pass, file free: the pass repairs the sheet before doing anything else.
  const { book: freeBook, results: freeResults, log: freeLog } = fakeBook();
  const outcomes = await pollInbox({
    db,
    book: freeBook,
    workflow,
    mailbox: fakeMailbox([]),
    pollMs: 1,
    timeoutMs: 50,
  });

  assert.equal(outcomes.length, 0, 'nothing new to do');
  assert.equal(freeResults.get(4), 'recorded ✓', 'the row now says it was recorded');
  assert.deepEqual(freeLog.map((entry) => entry.result), ['recorded ✓']);
  assert.equal(log.length, 0, 'the locked sheet was never written');
});

function fakeBook(
  rows: Record<string, number> = { 'steel bracket': 4 },
  options: { locked?: boolean; lockedWrites?: boolean; lockedClears?: boolean } = {},
) {
  const locked = Object.assign(new Error('EBUSY: resource busy or locked'), { code: 'EBUSY' });
  const guard = () => {
    if (options.locked) throw locked;
  };
  const guardWrites = () => {
    if (options.locked || options.lockedWrites) throw locked;
  };
  const drafts = new Map<number, string>();
  const marks = new Map<number, 'approved' | 'rejected'>();
  const results = new Map<number, string>();
  const log: { mail: string; did: string; answer: string; result: string }[] = [];

  const book: Workbook = {
    async ensureInterface() {},
    async products() {
      return Object.keys(rows).map((name) => ({ sku: name, name }));
    },
    async resolveRows(lines) {
      const matched = new Map<number, number>();
      const missing: string[] = [];
      lines.forEach((line, at) => {
        const row = rows[line.sku ?? line.description] ?? rows[line.description];
        if (row) matched.set(at, row);
        else missing.push(line.description);
      });
      return { rows: matched, missing };
    },
    async writeDraft(value, cell) {
      guard();
      drafts.set(cell.row, value);
    },
    async readApproval(row) {
      return marks.get(row) ?? 'pending';
    },
    async clearDraft(row) {
      // Measured 2026-10-28: this write happens inside the resumed step, so a lock here failed the
      // whole run and the human's rejection was lost.
      if (options.lockedClears) throw locked;
      drafts.delete(row);
    },
    async markResults(list) {
      guardWrites();
      for (const { row, text } of list) results.set(row, text);
    },
    async readResult(row) {
      return results.get(row) ?? '';
    },
    async appendLog(entries) {
      guardWrites();
      log.push(...entries.map(({ mail, did, answer, result }) => ({ mail, did, answer, result })));
    },
    async revertLastWrite() {
      drafts.clear();
    },
  };

  return { book, drafts, marks, results, log };
}

function auditLines(db: Db, runId: string): string[] {
  return db
    .prepare(`SELECT actor || ':' || action AS line FROM audit_log WHERE run_id = ? ORDER BY id`)
    .all(runId)
    .map((row) => (row as { line: string }).line);
}

function harness() {
  const db = openDb(':memory:');
  const { book, drafts, marks, results, log } = fakeBook();
  const makeWorkflow = (mastra = new Mastra({ storage: new InMemoryStore() })) =>
    buildWorkflow({
      mastra,
      book,
      extract: async () => EXTRACTION,
      provider: resolveProfile('openai'),
    });
  // One instance by default: an in-memory snapshot is only visible to the workflow that made it.
  const workflow = makeWorkflow();

  const pass = (options: { workflow?: ReturnType<typeof makeWorkflow>; pollMs?: number; timeoutMs?: number } = {}) =>
    pollInbox({
      db,
      book,
      workflow: options.workflow ?? workflow,
      mailbox: fakeMailbox(),
      pollMs: options.pollMs ?? 1,
      timeoutMs: options.timeoutMs ?? 50,
    });

  return { db, book, drafts, marks, results, log, makeWorkflow, pass };
}

test('a pass still runs when the workbook cannot be refreshed because Excel has it open', async () => {
  const { book, pass, drafts, marks } = harness();
  book.ensureInterface = async () => {
    throw Object.assign(new Error('EPERM: operation not permitted, rename'), { code: 'EPERM' });
  };
  marks.set(4, 'approved');

  const outcomes = await pass({ pollMs: 1, timeoutMs: 5_000 });

  assert.equal(outcomes[0].action, 'approved', 'reads and decisions still work while the file is open');
  assert.equal(drafts.get(4), '30 · steel bracket');
});

test('a workbook held open in Excel defers the mail instead of burning it', async () => {
  const db = openDb(':memory:');
  const { book } = fakeBook(undefined, { locked: true });
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  assert.equal(outcomes[0].action, 'deferred');
  assert.equal(getRun(db, MAIL.messageId), undefined, 'the ledger row is gone, so the next pass retries');
});

test('what an open workbook blocked still reaches the sheet, once it is free', async () => {
  const db = openDb(':memory:');
  const { book } = fakeBook(undefined, { locked: true });
  book.ensureInterface = async () => {
    throw Object.assign(new Error('EBUSY: resource busy or locked'), { code: 'EBUSY' });
  };
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  // Nothing could be written while Excel held the file — not even the explanation.
  assert.deepEqual(
    pendingNotices(db).map((row) => row.result),
    ['close Excel and it will appear', 'retried automatically'],
    'both the missing interface and the mail it blocked are queued',
  );

  // Excel is closed: the next pass puts both lines in the Log, in words the owner can act on.
  const { book: freeBook, log } = fakeBook();
  await pollInbox({ db, book: freeBook, workflow, mailbox: fakeMailbox([]), pollMs: 1, timeoutMs: 5 });

  assert.deepEqual(log.map((entry) => entry.result), [
    'close Excel and it will appear',
    'retried automatically',
  ]);
  assert.match(log[1].mail, /orders@acer\.example/, 'the line names the mail, not a Message-ID');
  assert.match(log[1].did, /was open in Excel/);

  // Said once. A file left open all afternoon must not fill the sheet with the same sentence.
  const { book: thirdBook, log: thirdLog } = fakeBook();
  await pollInbox({ db, book: thirdBook, workflow, mailbox: fakeMailbox([]), pollMs: 1, timeoutMs: 5 });
  assert.deepEqual(thirdLog, []);
  assert.equal(pendingNotices(db).length, 0);
});

test('a mail the model could not read is recorded where the owner looks', async () => {
  const db = openDb(':memory:');
  const { book, log } = fakeBook();
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => {
      throw new Error('schema: lines is required');
    },
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  assert.equal(outcomes[0].action, 'failed');
  // The workbook was writable, so it lands in the same pass — no terminal to read.
  assert.deepEqual(log.map((entry) => entry.result), ['not drafted']);
  assert.match(log[0].did, /needs a person/, 'the sheet says what to do, not what threw');
  const runId = getRun(db, MAIL.messageId)!.id;
  assert.ok(
    auditLines(db, runId).some((line) => /lines is required/.test(line)),
    'the technical reason is still in the audit trail',
  );
});

test('waiting survives the sheet being unreadable mid-poll', async () => {
  const { book, pass, marks } = harness();
  const read = book.readApproval.bind(book);
  let reads = 0;
  book.readApproval = async (row: number) => {
    reads += 1;
    // Excel is saving: the read fails, the driver must keep waiting rather than crash.
    if (reads === 2) throw Object.assign(new Error('EBUSY'), { code: 'EBUSY' });
    if (reads === 4) marks.set(4, 'approved');
    return read(row);
  };

  const outcomes = await pass({ pollMs: 1, timeoutMs: 5_000 });

  assert.equal(outcomes[0].action, 'approved');
});

test('a decision survives a write-back that Excel keeps blocking', async () => {
  const db = openDb(':memory:');
  // The human answered and closed the file, but something still holds it when we write back.
  const { book } = fakeBook(undefined, { lockedWrites: true });
  const answering: Workbook = { ...book, readApproval: async () => 'approved' as const };
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book: answering,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({
    db,
    book: answering,
    workflow,
    mailbox: fakeMailbox(),
    pollMs: 1,
    timeoutMs: 5_000,
    writeBack: { retries: 0, pauseMs: 1 },
  });

  assert.equal(outcomes[0].action, 'approved', 'the human\'s decision stands even though the sheet refused it');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'approved');
  const audited = db
    .prepare(`SELECT action FROM audit_log`)
    .all()
    .map((row) => (row as { action: string }).action);
  assert.ok(
    audited.some((action) => action.startsWith('sheet-write-failed')),
    'the ledger says why the sheet does not show it',
  );
});

test('a mail that was already handled is skipped without touching the model or the sheet', async () => {
  const { db, pass, drafts, marks } = harness();
  marks.set(4, 'approved');

  const first = await pass();
  assert.equal(first[0].action, 'approved');

  drafts.clear(); // prove the second pass wrote nothing
  const second = await pass();

  assert.equal(second[0].action, 'skipped');
  assert.equal(drafts.size, 0);
  const runs = db.prepare(`SELECT COUNT(*) AS n FROM workflow_runs`).get() as { n: number };
  assert.equal(runs.n, 1, 'one message, one run — however many polls');
});

test('the driver waits for a mark that arrives while it is already polling', async () => {
  const { book, marks, pass } = harness();
  const read = book.readApproval.bind(book);
  let reads = 0;
  book.readApproval = async (row: number) => {
    reads += 1;
    if (reads === 3) marks.set(4, 'approved'); // the human marks it mid-wait
    return read(row);
  };

  const outcomes = await pass({ pollMs: 1, timeoutMs: 5_000 });

  assert.equal(outcomes[0].action, 'approved');
  assert.ok(reads >= 3, 'the driver actually re-read the sheet');
});

test('an unmarked row leaves the run suspended, and the next pass finishes it', async () => {
  const { db, marks, pass } = harness();

  const first = await pass({ timeoutMs: 5 });
  assert.equal(first[0].action, 'pending');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'pending', 'silence is not an approval');
  assert.ok(
    auditLines(db, getRun(db, MAIL.messageId)!.id).some((line) => /left suspended/.test(line)),
  );

  marks.set(4, 'approved');
  const second = await pass({ timeoutMs: 5 });

  assert.equal(second[0].action, 'approved');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'approved');
  assert.deepEqual(
    auditLines(db, getRun(db, MAIL.messageId)!.id).filter((line) => line.startsWith('human:')),
    ['human:approved row 4'],
  );
});

test('a rejection is recorded against the human, row by row', async () => {
  const { db, pass, marks, drafts } = harness();
  marks.set(4, 'rejected');

  const outcomes = await pass();

  assert.equal(outcomes[0].action, 'rejected');
  const run = getRun(db, MAIL.messageId)!;
  assert.equal(run.status, 'rejected');
  assert.deepEqual(auditLines(db, run.id).filter((l) => l.startsWith('human:')), [
    'human:rejected row 4',
  ]);
  assert.equal(drafts.has(4), false, 'the rejected draft is cleared');
});

test('a mail whose SKUs are not in the sheet ends as unmatched, not as a rejection', async () => {
  const db = openDb(':memory:');
  const { book } = fakeBook({});
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  assert.equal(outcomes[0].action, 'unmatched');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'unmatched');
  assert.deepEqual(outcomes[0].unmatched, ['steel bracket']);
});

test('an extraction failure marks the run failed without waiting for a human', async () => {
  const db = openDb(':memory:');
  const { book } = fakeBook();
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => {
      throw Object.assign(new Error('could not parse'), { name: 'AI_NoObjectGeneratedError' });
    },
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  assert.equal(outcomes[0].action, 'failed');
  const run = getRun(db, MAIL.messageId)!;
  assert.equal(run.status, 'failed');
  const lines = auditLines(db, run.id);
  assert.equal(lines[0], 'pipeline:start-failed');
  assert.match(
    lines[1],
    /^pipeline:start-error: .*could not parse/,
    'the reason is kept for whoever has to debug it — the sheet only says "needs a person"',
  );
});

test('a workbook that is not a stock table is reported once, as a workbook problem', async () => {
  const db = openDb(':memory:');
  const { book, log } = fakeBook();
  // What the real client throws when neither a SKU nor a product-name column exists.
  book.products = async () => {
    throw Object.assign(new Error('stock.xlsx has no product column.'), { name: 'WorkbookShapeError' });
  };
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  const second = { ...MAIL, messageId: '<inbox-2@acer.example>' };
  const outcomes = await pollInbox({
    db,
    book,
    workflow,
    mailbox: fakeMailbox([MAIL, second]),
    pollMs: 1,
    timeoutMs: 5,
  });

  assert.deepEqual(outcomes.map((o) => o.action), ['failed', 'failed']);
  // Two mails, one line: the fix is the sheet, not the mail.
  assert.deepEqual(log.map((entry) => entry.result), ['fix the workbook']);
  assert.match(log[0].did, /has no product column/);
  assert.equal(log[0].mail, '(sheetflow itself)');
});

test('a mail that is not a purchase is recorded as ignored and writes nothing', async () => {
  const db = openDb(':memory:');
  const { book, drafts } = fakeBook();
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    // A pitch may mention products; "other" is the one label allowed to suppress a mail.
    extract: async () => ({ intent: 'other', lines: [{ description: 'steel bracket' }] }),
    provider: resolveProfile('openai'),
  });

  const outcomes = await pollInbox({
    db,
    book,
    workflow,
    mailbox: fakeMailbox(),
    pollMs: 1,
    timeoutMs: 5,
  });

  assert.equal(outcomes[0].action, 'ignored');
  assert.equal(getRun(db, MAIL.messageId)?.status, 'ignored');
  assert.equal(drafts.size, 0, 'a pitch never reaches the sheet');
  assert.ok(
    auditLines(db, getRun(db, MAIL.messageId)!.id).includes(
      'pipeline:intent=other; not a purchase, nothing drafted',
    ),
  );
});

test('quantities in the text with no lines extracted is flagged for a human', async () => {
  const db = openDb(':memory:');
  const { book } = fakeBook();
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    // What Nova Lite actually did to test/mail/10-injection-leak.eml: obeyed the injected
    // "SYSTEM NOTICE" far enough to return nothing, while nothing harmful happened.
    extract: async () => ({ intent: 'other' }),
    provider: resolveProfile('openai'),
  });

  await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  const run = getRun(db, MAIL.messageId)!;
  assert.equal(run.status, 'ignored');
  assert.ok(
    auditLines(db, run.id).includes(
      'pipeline:quantities in the text but no lines extracted — inspect this mail',
    ),
    auditLines(db, run.id).join('\n'),
  );
});

test('a rejection is written to the log, so the cleared draft leaves a history', async () => {
  const { pass, marks, drafts, log } = harness();
  marks.set(4, 'rejected');

  await pass();

  assert.equal(drafts.has(4), false, 'the working cell is cleared');
  assert.deepEqual(
    log.map(({ did, answer, result }) => ({ did, answer, result })),
    [{ did: 'row 4 · 30 · steel bracket', answer: 'no', result: 'cleared' }],
    'the log is the only place the history survives',
  );
});

test('a mail that never reached a row still leaves a log line', async () => {
  const db = openDb(':memory:');
  const { book, log } = fakeBook({});
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => EXTRACTION,
    provider: resolveProfile('openai'),
  });

  await pollInbox({ db, book, workflow, mailbox: fakeMailbox(), pollMs: 1, timeoutMs: 5 });

  assert.deepEqual(
    log.map(({ did, answer, result }) => ({ did, answer, result })),
    [{ did: 'no matching row: steel bracket', answer: '—', result: 'not in sheet' }],
  );
});

test('the human sees their answer recorded in their own row', async () => {
  const { pass, marks, results } = harness();
  marks.set(4, 'approved');

  await pass();

  // The sheet is the only interface: if typing OK produced no visible change, nobody would trust it.
  assert.equal(results.get(4), 'recorded ✓');
});

test('the audit log names the drafted rows before any human decision exists', async () => {
  const { db, pass, marks } = harness();
  marks.set(4, 'approved');

  await pass();

  const run = getRun(db, MAIL.messageId)!;
  assert.deepEqual(auditLines(db, run.id), [
    'pipeline:intent=order drafted rows [4] unmatched [] no-quantity []',
    'human:approved row 4',
  ]);
});




