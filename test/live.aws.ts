// Live check — NOT part of `npm test`. Spends real Bedrock tokens and needs AWS credentials.
//
//   npm run test:live
//
// The whole loop, against a throwaway copy of `test/inventory.xlsx` and the `.eml` in `test/mail`:
//   mail → Nova Lite → SKU rows → AI_Draft → (human marks the sheet) → approve/reject read back.

import assert from 'node:assert/strict';
import { copyFile, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { Mastra } from '@mastra/core';
import { InMemoryStore } from '@mastra/core/storage';
import { buildWorkflow } from '../src/workflow.js';
import { resolveProfile } from '../src/model.js';
import { parseMail } from '../src/mail.js';
import { fileWorkbook } from '../src/excel.js';
import { readFile } from 'node:fs/promises';

const profile = process.env.LLM_PROFILE || 'aws';
const provider = resolveProfile(profile);

const dir = await mkdtemp(join(tmpdir(), 'sheetflow-live-'));
const bookPath = join(dir, 'inventory.xlsx');
await copyFile('test/inventory.xlsx', bookPath);

const mail = await parseMail(await readFile('test/mail/01-all-in-stock.eml'));
const book = fileWorkbook(bookPath);

const workflow = buildWorkflow({
  provider,
  book,
  mastra: new Mastra({ storage: new InMemoryStore() }),
});

const startedAt = Date.now();
const run = await workflow.createRun();
const first = await run.start({
  inputData: {
    messageId: mail.messageId,
    from: mail.from,
    subject: mail.subject,
    body: mail.body,
    receivedAt: mail.receivedAt,
  },
});

console.log(`profile=${profile} model=${provider.model} elapsed=${Date.now() - startedAt}ms`);
console.log(`status=${first.status}`);

assert.equal(first.status, 'suspended', `expected a human checkpoint, got ${first.status}`);

// Read the sheet the way a human would see it.
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(bookPath);
const ws = wb.worksheets[0];
const headers = ws.getRow(1).values as unknown[];
const draftCol = headers.indexOf('AI suggests');
const reviewCol = headers.indexOf('Your answer: OK or no');
assert.ok(draftCol > 0 && reviewCol > 0, 'the AI columns were created');

const drafted: { row: number; sku: string; draft: unknown }[] = [];
for (let row = 2; row <= ws.rowCount; row++) {
  const value = ws.getCell(row, draftCol).value;
  if (value === null || value === undefined) continue;
  drafted.push({ row, sku: String(ws.getCell(row, 1).value), draft: value });
}
console.log('drafted:', JSON.stringify(drafted));

assert.ok(drafted.length >= 2, `expected both SKUs drafted, got ${drafted.length}`);
assert.deepEqual(
  drafted.map((d) => d.sku),
  ['AP-TSH-BLK-M', 'ACC-CAP-NVY-OS'],
);

// The human decides: one row OK, the other rejected.
const [approved, rejected] = drafted;
ws.getCell(approved.row, reviewCol).value = 'OK';
ws.getCell(rejected.row, reviewCol).value = 'no';
await wb.xlsx.writeFile(bookPath);

// Exactly what `npm run inbox` sends: the answers read off the sheet, as data.
const resumed = await run.resume({
  step: 'review',
  resumeData: {
    decisions: [
      { row: approved.row, decision: 'approved' as const },
      { row: rejected.row, decision: 'rejected' as const },
    ],
  },
});
console.log(`resume=${resumed.status}`, JSON.stringify(resumed.status === 'success' ? resumed.result : resumed));

if (resumed.status !== 'success') assert.fail(`expected success, got ${resumed.status}`);
assert.deepEqual(resumed.result.approved, [approved.row]);
assert.deepEqual(resumed.result.rejected, [rejected.row]);

// The rejected suggestion is cleared by the driver's write-back, not by the graph; that is covered by
// `test/inbox.test.ts`. This script only proves the live model path and the resume contract.
const after = new ExcelJS.Workbook();
await after.xlsx.readFile(bookPath);
const sheet = after.worksheets[0];
assert.notEqual(sheet.getCell(approved.row, draftCol).value, null);

console.log(`ok — approved row ${approved.row}, rejected row ${rejected.row}, sheet ${bookPath}`);
