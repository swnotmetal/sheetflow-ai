// The pipeline test: mail -> extract -> SKU rows -> AI_Draft -> human mark in the sheet.
//
// Deterministic by construction — the model is a mock, the workbook is memory. What is *not*
// mocked is the graph: real Mastra steps, real suspend/resume, real approval reading.

import assert from 'node:assert/strict';
import test from 'node:test';
import { Mastra } from '@mastra/core';
import { InMemoryStore } from '@mastra/core/storage';
import { buildWorkflow, decideIntent, formatDraft, groundText } from '../src/workflow.js';
import { resolveProfile, type Extraction } from '../src/model.js';
import type { Workbook } from '../src/excel.js';

const MAIL = {
  messageId: '<MSG-1@acer.example>',
  from: 'orders@acer.example',
  subject: 'PO 4500 — restock request',
  body: 'Please ship 30 units of steel bracket to our depot in Rotterdam.',
};

const EXTRACTION: Extraction = {
  intent: 'order',
  customer: ['Acer Industrial'],
  po: ['4500'],
  lines: [{ description: 'steel bracket', quantity: '30' }],
};

interface FakeSheet {
  /** Line text → sheet row, as the sheet would answer. */
  rows?: Record<string, number>;
}

function memoryBook({ rows: sheet = { 'steel bracket': 4 } }: FakeSheet = {}) {
  const drafts = new Map<number, string>();
  const review = new Map<number, 'approved' | 'rejected' | 'pending'>();
  const cleared: number[] = [];
  const results = new Map<number, string>();
  const log: { did: string; answer: string; result: string }[] = [];

  const book: Workbook = {
    async ensureInterface() {},
    async products() {
      return Object.keys(sheet).map((name) => ({ sku: name, name }));
    },
    async resolveRows(lines) {
      const rows = new Map<number, number>();
      const missing: string[] = [];
      lines.forEach((line, at) => {
        const row = sheet[line.sku ?? line.description] ?? sheet[line.description];
        if (row) rows.set(at, row);
        else missing.push(line.description);
      });
      return { rows, missing };
    },
    async writeDraft(value, cell) {
      drafts.set(cell.row, value);
    },
    async readApproval(row) {
      return review.get(row) ?? 'pending';
    },
    async clearDraft(row) {
      drafts.delete(row);
      cleared.push(row);
    },
    async readResult(row: number) {
      return results.get(row) ?? '';
    },
    async markResults(list) {
      for (const { row, text } of list) results.set(row, text);
    },
    async appendLog(entries) {
      log.push(...entries);
    },
    async revertLastWrite() {
      drafts.clear();
    },
  };

  return { book, drafts, review, cleared, results, log };
}

function harness(extraction: Extraction = EXTRACTION, sheet: FakeSheet = {}) {
  const fake = memoryBook(sheet);
  const workflow = buildWorkflow({
    // Its own snapshot store: one test's suspended run can never be resumed by another.
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book: fake.book,
    extract: async () => extraction,
    // Named explicitly: nothing selects a model implicitly, and no network is touched here.
    provider: resolveProfile('openai'),
  });
  return { workflow, ...fake };
}

test('the row comes from the sheet, not from the caller', async () => {
  const { workflow, drafts } = harness();
  const run = await workflow.createRun();

  const first = await run.start({ inputData: MAIL });

  assert.equal(first.status, 'suspended');
  assert.deepEqual([...drafts.entries()], [[4, '30']], 'drafted into the row the SKU lives in');
});

test('a line the sheet has no row for is reported, never written', async () => {
  const { workflow, drafts } = harness(
    {
      intent: 'order',
      lines: [
        { description: 'steel bracket', quantity: '30' },
        { description: 'unobtainium', quantity: '5' },
      ],
    },
    { rows: { 'steel bracket': 4 } },
  );
  const run = await workflow.createRun();

  await run.start({ inputData: MAIL });

  assert.deepEqual([...drafts.keys()], [4], 'only the matched SKU got a cell');
});

test('a line with no quantity is drafted as `?` — a guessed number is never written', async () => {
  const { workflow, drafts } = harness({ intent: 'order', lines: [{ description: 'steel bracket' }] });
  const run = await workflow.createRun();

  await run.start({ inputData: MAIL });

  assert.deepEqual([...drafts.values()], ['?']);
});

test('a quantity in the text wins over a model that called it a question', async () => {
  const { workflow, drafts } = harness({
    intent: 'question',
    lines: [{ description: 'steel bracket', quantity: '30' }],
  });
  const run = await workflow.createRun();

  await run.start({ inputData: MAIL });

  // Nova Lite called three plainly-quantified orders "question" while these fixtures were being
  // written; believing it would silently drop real orders.
  assert.deepEqual([...drafts.values()], ['30']);
});

test('"do you have any black crew socks?" still reaches the sheet', async () => {
  const { workflow, drafts } = harness({
    intent: 'question',
    lines: [{ description: 'steel bracket' }],
  });
  const run = await workflow.createRun();

  const first = await run.start({ inputData: MAIL });

  // A bare availability question is not ours to interpret: the human sees the row and answers.
  assert.equal(first.status, 'suspended');
  assert.deepEqual([...drafts.values()], ['?']);
});

test('a vendor pitch is not drafted either', async () => {
  const { workflow, drafts } = harness({ intent: 'other', lines: [] });
  const run = await workflow.createRun();

  await run.start({ inputData: MAIL });

  assert.equal(drafts.size, 0);
});

test('intent is decided by evidence, not by the label alone', () => {
  assert.equal(decideIntent('question', [{ quantity: '30' }]), 'order');
  assert.equal(decideIntent('order', [{ quantity: '30' }]), 'order');
  assert.equal(decideIntent('order', [{ quantity: null }]), 'question', 'no quantity');
  assert.equal(decideIntent('other', []), 'other');
  assert.equal(decideIntent('order', []), 'question');
});

test('the subject is part of the grounding text, not ignored', async () => {
  const { workflow } = harness({ intent: 'order', po: ['4500'], lines: [] });
  const run = await workflow.createRun();

  // A PO number that appears only in the subject must survive grounding.
  const first = await run.start({ inputData: MAIL });
  assert.equal(first.status, 'suspended');
  assert.match(groundText(MAIL), /PO 4500/);
});

test('the human answers; the answer arrives as the resume payload', async () => {
  const { workflow, drafts } = harness();
  const run = await workflow.createRun();
  await run.start({ inputData: MAIL });

  // The driver reads the sheet and hands the answer over — the step itself reads no file.
  const resumed = await run.resume({ step: 'review', resumeData: { decisions: [{ row: 4, decision: 'approved' }] } });

  if (resumed.status !== 'success') assert.fail(`expected success, got ${resumed.status}`);
  assert.deepEqual(resumed.result.approved, [4]);
  assert.deepEqual(resumed.result.pending, []);
  assert.equal(drafts.get(4), '30', 'an approval leaves the draft in place for the human to act on');
});

test('a rejection is recorded; the draft is cleared later, by the write-back', async () => {
  const { workflow, drafts } = harness();
  const run = await workflow.createRun();
  await run.start({ inputData: MAIL });

  const resumed = await run.resume({ step: 'review', resumeData: { decisions: [{ row: 4, decision: 'rejected' }] } });

  if (resumed.status !== 'success') assert.fail(`expected success, got ${resumed.status}`);
  assert.deepEqual(resumed.result.rejected, [4]);
  assert.equal(drafts.get(4), '30', 'the graph only decides; the driver does the clearing');
});

test('an untouched row stays pending — silence is not consent', async () => {
  const { workflow } = harness();
  const run = await workflow.createRun();
  await run.start({ inputData: MAIL });

  const resumed = await run.resume({ step: 'review', resumeData: { decisions: [] } });

  if (resumed.status !== 'success') assert.fail(`expected success, got ${resumed.status}`);
  assert.deepEqual(resumed.result.pending, [4]);
  assert.deepEqual(resumed.result.approved, []);
  assert.deepEqual(resumed.result.rejected, []);
});

test('a resume needs no file access — the answer is the input', async () => {
  const { workflow, book } = harness();
  const run = await workflow.createRun();
  await run.start({ inputData: MAIL });

  // Excel is holding the workbook. The graph must not care: it is given the answer, not asked for it.
  book.readApproval = async () => {
    throw Object.assign(new Error('EBUSY: resource busy or locked'), { code: 'EBUSY' });
  };

  const resumed = await run.resume({ step: 'review', resumeData: { decisions: [{ row: 4, decision: 'approved' }] } });

  if (resumed.status !== 'success') assert.fail(`expected success, got ${resumed.status}`);
  assert.deepEqual(resumed.result.approved, [4]);
});

test('a failed extraction fails the run instead of drafting an empty row', async () => {
  const { book } = memoryBook();
  const workflow = buildWorkflow({
    mastra: new Mastra({ storage: new InMemoryStore() }),
    book,
    extract: async () => {
      throw Object.assign(new Error('could not parse'), { name: 'AI_NoObjectGeneratedError' });
    },
    provider: resolveProfile('openai'),
  });
  const run = await workflow.createRun();

  const result = await run.start({ inputData: MAIL });

  assert.equal(result.status, 'failed');
});

test('formatDraft keeps the human-facing value readable', () => {
  assert.equal(formatDraft([{ description: 'steel bracket', quantity: '30' }]), '30 steel bracket');
  assert.equal(
    formatDraft([{ description: 'steel bracket', quantity: '2', unit: 'box' }]),
    '2 box steel bracket',
  );
  assert.equal(formatDraft([]), '');
});
