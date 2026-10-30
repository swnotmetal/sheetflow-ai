// The Google Sheets client, exercised against a fake spreadsheet.
//
// No network and no Google project: the transport seam is a map of tabs, which is exactly what the
// REST client turns a spreadsheets resource into. What is asserted is that the *same* rules hold as
// on a local file — one header discovery, exact matching, the three columns, the Log tab — because a
// shop's data must not behave differently depending on where the sheet lives.

import assert from 'node:assert/strict';
import test from 'node:test';
import { generateKeyPairSync, verify } from 'node:crypto';
import { Mastra } from '@mastra/core';
import { InMemoryStore } from '@mastra/core/storage';
import { columnName, quoteSheet, sheetWorkbook, signedAssertion } from '../src/sheets.js';
import type { SheetsTransport } from '../src/sheets.js';
import { buildWorkflow } from '../src/workflow.js';
import { resolveProfile } from '../src/model.js';

interface Fake extends SheetsTransport {
  tabs: Map<string, string[][]>;
  notes: Map<string, string>;
}

/** An in-memory spreadsheet: `tabs` holds rows, and cells are addressed the way a human reads them. */
function fakeSheets(initial: Record<string, string[][]> = {}): Fake {
  const tabs = new Map(Object.entries(initial).map(([name, rows]) => [name, rows.map((row) => [...row])]));
  const notes = new Map<string, string>();
  const put = (sheet: string, row: number, col: number, text: string) => {
    const grid = tabs.get(sheet)!;
    while (grid.length < row) grid.push([]);
    const line = grid[row - 1];
    while (line.length < col) line.push('');
    line[col - 1] = text;
  };

  return {
    tabs,
    notes,
    async sheets() {
      let id = 0;
      return [...tabs.keys()].map((name) => ({ name, id: id++ }));
    },
    async grid(sheet) {
      return (tabs.get(sheet) ?? []).map((row) => [...row]);
    },
    async cell(sheet, row, col) {
      return tabs.get(sheet)?.[row - 1]?.[col - 1] ?? '';
    },
    async write(sheet, row, col, text) {
      put(sheet, row, col, text);
    },
    async note(sheet, row, col, text) {
      const key = `${sheet}!${columnName(col)}${row}`;
      if (text === null) notes.delete(key);
      else notes.set(key, text);
    },
    async append(sheet, values) {
      while (!tabs.has(sheet)) tabs.set(sheet, []);
      tabs.get(sheet)!.push([...values]);
    },
    async addSheet(name) {
      if (!tabs.has(name)) tabs.set(name, []);
    },
  };
}

const STOCK = [
  ['SKU', 'Product Name', 'Qty'],
  ['AP-TSH-BLK-M', 'Oversized Heavyweight Tee - Black - M', '10'],
  ['ACC-SOC-BLK-OS', 'Ribbed Athletic Crew Socks - Black', '5'],
];

test('a sheet is addressed by letter, the way a human reads it', () => {
  assert.equal(columnName(1), 'A');
  assert.equal(columnName(12), 'L');
  assert.equal(columnName(26), 'Z');
  assert.equal(columnName(27), 'AA');
  assert.equal(columnName(702), 'ZZ');
  assert.equal(columnName(703), 'AAA');
});

test('a tab name with an apostrophe is quoted the way the UI shows it', () => {
  assert.equal(quoteSheet('Sheet1'), "'Sheet1'");
  assert.equal(quoteSheet("Sam's stock"), "'Sam''s stock'");
});

test('the catalogue and the row matching work exactly as on a local file', async () => {
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  assert.deepEqual(await book.products(), [
    { sku: 'AP-TSH-BLK-M', name: 'Oversized Heavyweight Tee - Black - M' },
    { sku: 'ACC-SOC-BLK-OS', name: 'Ribbed Athletic Crew Socks - Black' },
  ]);

  const { rows, missing } = await book.resolveRows([
    { description: 'ribbed athletic crew socks - black' },
    { description: 'made up', sku: 'MADE-UP-1' },
  ]);
  assert.deepEqual([...rows], [[0, 3]]);
  assert.deepEqual(missing, ['made up']);
});

test('a draft lands in "AI suggests" with a note naming the mail', async () => {
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  await book.writeDraft('48 · Oversized Heavyweight Tee - Black - M', {
    row: 2,
    note: '2026-10-30 07:12 UTC · orders@acme.example — "Re: restock"',
  });

  assert.equal(transport.tabs.get('Stock')![0][3], 'AI suggests');
  assert.equal(transport.tabs.get('Stock')![0][4], 'Your answer: OK or no');
  assert.equal(transport.tabs.get('Stock')![0][5], 'AI result');
  assert.equal(transport.tabs.get('Stock')![1][3], '48 · Oversized Heavyweight Tee - Black - M');
  assert.match(transport.notes.get('Stock!D2') ?? '', /orders@acme\.example/);
});

test('the human answer decides, and the answer cell starts empty', async () => {
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  await book.writeDraft('48', { row: 2 });
  // A previous `OK` must not survive a re-draft: that would be consent for the previous number.
  transport.tabs.get('Stock')![1][4] = 'OK';
  await book.writeDraft('45', { row: 2 });
  assert.equal(await book.readApproval(2), 'pending');

  transport.tabs.get('Stock')![1][4] = 'ok';
  assert.equal(await book.readApproval(2), 'approved');
  transport.tabs.get('Stock')![1][4] = 'no';
  assert.equal(await book.readApproval(2), 'rejected');
});

test('a rejection clears the suggestion and its note', async () => {
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  await book.writeDraft('48', { row: 2, note: 'a mail' });
  await book.clearDraft(2);

  assert.equal(transport.tabs.get('Stock')![1][3], '');
  assert.equal(transport.notes.has('Stock!D2'), false);
});

test('the outcome is written where the human looked', async () => {
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });
  await book.writeDraft('48', { row: 2 });

  await book.markResults([{ row: 2, text: 'recorded ✓' }]);

  assert.equal(await book.readResult(2), 'recorded ✓');
  assert.equal(transport.tabs.get('Stock')![1][3], '48', 'the suggestion itself is not rewritten');
});

test('history goes to a Log tab of our own, never into one the shop already has', async () => {
  const transport = fakeSheets({
    Stock: STOCK,
    Log: [['my own notes']],
  });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  await book.appendLog([
    { when: '2026-10-30 09:00', mail: 'a mail', did: 'row 2 · 48', answer: 'OK', result: 'recorded ✓' },
  ]);

  assert.deepEqual(transport.tabs.get('Log'), [['my own notes']], 'their tab is untouched');
  assert.deepEqual(transport.tabs.get('Log (sheetflow)')![0], [
    'When (UTC)',
    'Mail',
    'What the AI did',
    'Your answer',
    'Result',
  ]);
  assert.equal(transport.tabs.get('Log (sheetflow)')![1][3], 'OK');

  // Said once: a second pass appends to the same tab instead of making another.
  await book.appendLog([
    { when: '2026-10-30 09:05', mail: 'b mail', did: 'row 2', answer: 'no', result: 'cleared' },
  ]);
  assert.equal(transport.tabs.get('Log (sheetflow)')!.length, 3);
});

test('the interface appears before any mail, and only once', async () => {
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  await book.ensureInterface();
  assert.deepEqual(transport.tabs.get('Stock')![0].slice(3), [
    'AI suggests',
    'Your answer: OK or no',
    'AI result',
  ]);
  assert.deepEqual(transport.tabs.get('Log')![0], [
    'When (UTC)',
    'Mail',
    'What the AI did',
    'Your answer',
    'Result',
  ]);

  await book.ensureInterface();
  assert.equal(transport.tabs.get('Stock')![0].length, 6, 'a second pass adds no fourth column');
  assert.equal(transport.tabs.size, 2, 'and no second Log tab');
});

test('a sheet with no product column says so, as on a local file', async () => {
  const transport = fakeSheets({ Stock: [['Ref', 'Qty'], ['A-1', '10']] });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  await assert.rejects(book.products(), /no product column/);
  await assert.rejects(book.resolveRows([{ description: 'x' }]), /no product column/);
});

test('a name-only sheet is matched by name', async () => {
  const transport = fakeSheets({
    Stock: [['Product Name', 'Qty'], ['Ribbed Athletic Crew Socks - Black', '5']],
  });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });

  assert.deepEqual(await book.products(), [{ sku: '', name: 'Ribbed Athletic Crew Socks - Black' }]);
  assert.deepEqual([...(await book.resolveRows([{ description: 'ribbed athletic crew socks - black' }])).rows], [[0, 2]]);
});

test('the real graph drafts into a Google Sheet, unchanged', async () => {
  // The point of the seam: nothing in the graph, the ledger or the waiting knows which spreadsheet
  // it is talking to. Same steps, same suspend, same three columns.
  const transport = fakeSheets({ Stock: STOCK });
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport });
  const workflow = buildWorkflow({
    book,
    mastra: new Mastra({ storage: new InMemoryStore() }),
    provider: resolveProfile('openai'),
    extract: async () => ({
      intent: 'order',
      lines: [{ description: 'Ribbed Athletic Crew Socks - Black', quantity: '12' }],
    }),
  });

  const run = await workflow.createRun();
  const result = await run.start({
    inputData: {
      messageId: '<g@acme.example>',
      from: 'ACME <orders@acme.example>',
      subject: 'Re: socks',
      body: 'Please send 12 of Ribbed Athletic Crew Socks - Black.',
      receivedAt: '2026-10-30T07:12:34.000Z',
    },
  });

  assert.equal(result.status, 'suspended', 'a human still decides, on any backend');
  assert.equal(transport.tabs.get('Stock')![2][3], '12 · Ribbed Athletic Crew Socks - Black');
  assert.match(transport.notes.get('Stock!D3') ?? '', /orders@acme\.example/);
  assert.equal(transport.tabs.get('Stock')![2][4], '', 'the answer cell is empty, waiting');

  // The human types OK in the browser; the answer is read exactly as it would be from a file.
  transport.tabs.get('Stock')![2][4] = 'OK';
  assert.equal(await book.readApproval(3), 'approved');
});

test('reverting refuses, and names where undo actually lives', async () => {
  const book = sheetWorkbook({ spreadsheetId: 'abc', transport: fakeSheets({ Stock: STOCK }) });
  await assert.rejects(book.revertLastWrite(), /Version history/);
});

test('the service-account assertion is a real RS256 signature of its own claims', () => {
  const { privateKey, publicKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
  const account = {
    client_email: 'sheetflow@example.iam.gserviceaccount.com',
    private_key: privateKey.export({ type: 'pkcs8', format: 'pem' }).toString(),
  };

  const jwt = signedAssertion(account, 1_700_000_000);
  const [header, claim, signature] = jwt.split('.');

  assert.deepEqual(JSON.parse(Buffer.from(header, 'base64url').toString()), {
    alg: 'RS256',
    typ: 'JWT',
  });
  const claims = JSON.parse(Buffer.from(claim, 'base64url').toString());
  assert.equal(claims.iss, account.client_email);
  assert.equal(claims.aud, 'https://oauth2.googleapis.com/token');
  assert.equal(claims.exp - claims.iat, 3600);

  // The signature must verify against the public key: a token Google would reject is worthless.
  const input = `${header}.${claim}`;
  const ok = verify('RSA-SHA256', Buffer.from(input), publicKey, Buffer.from(signature, 'base64url'));
  assert.equal(ok, true);
});
