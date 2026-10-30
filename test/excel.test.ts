// The real `.xlsx` client, exercised against a throwaway copy of the sample workbook.
//
// The assertions are about the guarantees: the catalogue comes from the sheet, a proposed SKU is
// verified against it, existing formulas survive, the human's mark decides, and the human can see
// that their answer was recorded.

import assert from 'node:assert/strict';
import test from 'node:test';
import { copyFile, mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { cellTextValue, fileWorkbook, isSheetBusy, parseDecision } from '../src/excel.js';

const SOURCE = 'test/inventory.xlsx';

async function scratchCopy() {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-'));
  const path = join(dir, 'inventory.xlsx');
  await copyFile(SOURCE, path);
  return path;
}

/** A throwaway workbook built to order, for the shapes the sample does not have. */
async function buildWorkbook(fill: (wb: ExcelJS.Workbook) => void) {
  const wb = new ExcelJS.Workbook();
  fill(wb);
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-'));
  const path = join(dir, 'workbook.xlsx');
  await wb.xlsx.writeFile(path);
  return path;
}

async function readCell(path: string, ref: string) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  return wb.worksheets[0].getCell(ref).value;
}

/** exceljs types `row.values` as an array *or* a map; these tests only ever read the array form. */
function rowValues(row: ExcelJS.Row): ExcelJS.CellValue[] {
  return row.values as ExcelJS.CellValue[];
}

async function readSheet(path: string) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  return wb.worksheets[0];
}

async function readSheetNamed(path: string, name: string) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  return wb.getWorksheet(name);
}

/** `solid` is the highlight that means "this row is asking for an answer"; anything else means quiet. */
function highlighted(cell: ExcelJS.Cell): boolean {
  return (cell.fill as { pattern?: string } | undefined)?.pattern === 'solid';
}

async function writeCell(path: string, ref: string, value: string) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  wb.worksheets[0].getCell(ref).value = value;
  await wb.xlsx.writeFile(path);
}

test('the catalogue is read from the sheet itself', async () => {
  const products = await fileWorkbook(await scratchCopy()).products();

  assert.equal(products.length, 27);
  assert.deepEqual(products[0], {
    sku: 'AP-TSH-BLK-M',
    name: 'Oversized Heavyweight Tee - Black - M',
  });
  assert.ok(products.some((product) => product.sku === 'ACC-SOC-BLK-OS'));
});

test('a line resolves by its SKU, by its SKU cell, or by its product name', async () => {
  const wb = fileWorkbook(await scratchCopy());

  const { rows, missing } = await wb.resolveRows([
    { description: 'ap-tsh-blk-m' }, // the description is already the SKU
    { description: 'some black socks', sku: 'ACC-SOC-BLK-OS' }, // the model resolved it
    { description: 'Oversized Heavyweight Tee - White - M' }, // full product name
    { description: 'ACC-NOT-A-SKU' },
  ]);

  assert.equal(rows.get(0), 2);
  assert.equal(rows.get(1), 19);
  assert.equal(rows.get(2), 5);
  assert.deepEqual(missing, ['ACC-NOT-A-SKU']);
});

test('a made-up SKU is never trusted over the sheet', async () => {
  const wb = fileWorkbook(await scratchCopy());

  const { rows, missing } = await wb.resolveRows([
    { description: 'Ribbed Athletic Crew Socks - Black', sku: 'ACC-SOC-BLK-XXL' },
  ]);

  // The invented SKU does not exist; the exact product name still resolves.
  assert.equal(rows.get(0), 19);
  assert.deepEqual(missing, []);

  const onlyInvented = await wb.resolveRows([{ description: 'nonsense', sku: 'MADE-UP-1' }]);
  assert.equal(onlyInvented.rows.size, 0);
  assert.deepEqual(onlyInvented.missing, ['nonsense']);
});

test('a draft lands in "AI suggests" and existing formulas are untouched', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);

  await wb.writeDraft('30', { row: 2 });

  assert.equal(await readCell(path, 'L1'), 'AI suggests');
  assert.equal(await readCell(path, 'M1'), 'Your answer: OK or no');
  assert.equal(await readCell(path, 'N1'), 'AI result');
  assert.equal(await readCell(path, 'L2'), '30');
  assert.equal(await readCell(path, 'A2'), 'AP-TSH-BLK-M', 'business data is untouched');
  assert.deepEqual(
    await readCell(path, 'I2'),
    { formula: 'F2+G2-H2', result: 30 },
    'the formula is still a formula',
  );
});

test('every flush keeps a .bak of the previous file', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);

  await wb.writeDraft('30', { row: 2 });

  const backup = await new ExcelJS.Workbook().xlsx.readFile(path + '.bak');
  assert.equal(backup.worksheets[0].getCell('L2').value, null, 'the backup predates the draft');
});

test('the human mark decides: approved, rejected, or still pending', async () => {
  const path = await scratchCopy();
  await fileWorkbook(path).writeDraft('30', { row: 2 });

  assert.equal(await fileWorkbook(path).readApproval(2), 'pending');

  await writeCell(path, 'M2', 'OK');
  assert.equal(await fileWorkbook(path).readApproval(2), 'approved');

  await writeCell(path, 'M2', 'no');
  assert.equal(await fileWorkbook(path).readApproval(2), 'rejected');
});

test('the outcome is written back where the human looked', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);
  await wb.writeDraft('30', { row: 2 });
  await wb.writeDraft('12', { row: 13 });

  await wb.markResults([
    { row: 2, text: 'recorded ✓' },
    { row: 13, text: 'cleared' },
  ]);

  assert.equal(await readCell(path, 'N2'), 'recorded ✓');
  assert.equal(await readCell(path, 'N13'), 'cleared');
  assert.equal(await readCell(path, 'L2'), '30', 'the suggestion itself is not rewritten');
});

test('a rejected draft can be cleared, and the last write reverted', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);

  await wb.writeDraft('30', { row: 2 });
  await wb.clearDraft(2);
  assert.equal(await readCell(path, 'L2'), null);

  await wb.writeDraft('99', { row: 2 });
  await wb.revertLastWrite();
  assert.equal(await readCell(path, 'L2'), null, 'revert restored the last committed file');
});

test('a suggestion written before the highlight existed is backfilled', async () => {
  const path = await scratchCopy();
  await fileWorkbook(path).ensureInterface();
  // Simulate a workbook from before this highlight: a value in `AI suggests`, no formatting.
  await writeCell(path, 'L12', '40');
  await writeCell(path, 'M13', 'OK');

  await fileWorkbook(path).ensureInterface();
  const sheet = await readSheet(path);

  assert.equal(sheet.getCell('M12').dataValidation?.type, 'list', 'the suggested row asks');
  assert.equal(highlighted(sheet.getCell('L12')), true);
  assert.equal(sheet.getCell('M13').dataValidation, undefined, 'a row without a suggestion stays quiet');
});

test('only a row carrying a suggestion asks for an answer', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);
  await wb.ensureInterface();

  // Nothing drafted yet: no picker anywhere, so nothing invites a meaningless answer.
  const before = await readSheet(path);
  assert.equal(before.getCell('M2').dataValidation, undefined);
  assert.equal(highlighted(before.getCell('L2')), false);

  await wb.writeDraft('30', { row: 2 });
  const drafted = await readSheet(path);
  assert.equal(drafted.getCell('L2').value, '30');
  assert.equal(drafted.getCell('M2').value, null, 'the answer starts empty');
  assert.equal(drafted.getCell('M2').dataValidation?.type, 'list');
  assert.equal(highlighted(drafted.getCell('L2')), true, 'the row asking for an answer is visible');
  assert.equal(highlighted(drafted.getCell('M2')), true);
  assert.equal(drafted.getCell('M3').dataValidation, undefined, 'untouched rows stay quiet');

  // A second suggestion on the same row must not inherit the first answer.
  await writeCell(path, 'M2', 'OK');
  await wb.writeDraft('45', { row: 2 });
  assert.equal(await readCell(path, 'M2'), null);
  assert.equal(await readCell(path, 'L2'), '45');

  // Answered: the row stops asking and keeps the receipt.
  await wb.markResults([{ row: 2, text: 'recorded ✓' }]);
  const done = await readSheet(path);
  assert.equal(done.getCell('M2').dataValidation, undefined);
  assert.equal(highlighted(done.getCell('L2')), false);
  assert.equal(done.getCell('N2').value, 'recorded ✓');
});

test('the interface appears before any mail, and only once', async () => {
  const path = await scratchCopy();

  await fileWorkbook(path).ensureInterface();

  const first = await new ExcelJS.Workbook().xlsx.readFile(path);
  assert.deepEqual(rowValues(first.worksheets[0].getRow(1)).slice(12), [
    'AI suggests',
    'Your answer: OK or no',
    'AI result',
  ]);
  assert.deepEqual(rowValues(first.getWorksheet('Log')!.getRow(1)).slice(1), [
    'When (UTC)',
    'Mail',
    'What the AI did',
    'Your answer',
    'Result',
  ]);

  // Idempotent: a second pass must not add a fourth column or a second Log sheet.
  await fileWorkbook(path).ensureInterface();
  const second = await new ExcelJS.Workbook().xlsx.readFile(path);
  assert.equal(second.worksheets[0].columnCount, 14);
  assert.equal(second.worksheets.length, 2);
});

test('the log sheet keeps history a cleared draft would lose', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);
  await wb.writeDraft('30', { row: 2 });
  await wb.clearDraft(2);

  await wb.appendLog([
    { when: '2026-10-28 09:00', mail: '<m@x>', did: 'row 2 · AP-TSH-BLK-M · 30', answer: 'no', result: 'cleared' },
    { when: '2026-10-28 09:05', mail: '<n@x>', did: 'not a purchase', answer: '—', result: 'ignored' },
  ]);

  assert.equal(await readCell(path, 'L2'), null, 'the working cell is gone');
  const sheet = (await new ExcelJS.Workbook().xlsx.readFile(path)).getWorksheet('Log')!;
  assert.equal(sheet.rowCount, 3);
  assert.deepEqual(rowValues(sheet.getRow(2)).slice(1), [
    '2026-10-28 09:00',
    '<m@x>',
    'row 2 · AP-TSH-BLK-M · 30',
    'no',
    'cleared',
  ]);
});

test('the sample sheet is exactly as it was before any of this', async () => {
  // Guard: the tests above work on copies, so the fixture itself must never be rewritten.
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SOURCE);
  assert.equal(wb.worksheets[0].columnCount, 11);
  assert.equal(wb.worksheets[0].getCell('L1').value, null);
});

test('the suggestion carries the mail that asked for it, as a note', async () => {
  const path = await scratchCopy();
  const wb = fileWorkbook(path);
  const first = '2026-10-30 09:12 UTC · orders@acme.example — "Re: restock"';

  await wb.writeDraft('30 · Oversized Heavyweight Tee - Black - M', { row: 2, note: first });

  const sheet = await readSheet(path);
  assert.equal(sheet.getCell('L2').note, first, 'when, who and the subject sit on the row');
  assert.equal(sheet.getCell('A2').note, undefined, 'business cells are not annotated');

  // A second mail on the same row must not leave the first sender behind.
  await wb.writeDraft('45 · Oversized Heavyweight Tee - Black - M', { row: 2, note: 'second mail' });
  assert.equal((await readSheet(path)).getCell('L2').note, 'second mail');

  // Rejecting returns the row to plain stock data: the provenance goes with the suggestion.
  await wb.clearDraft(2);
  const cleared = await readSheet(path);
  assert.equal(cleared.getCell('L2').value, null);
  assert.equal(cleared.getCell('L2').note, undefined);
});

test('a workbook with no SKU column is matched by product name', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Stock');
    ws.addRow(['Product Name', 'Qty']);
    ws.addRow(['Oversized Heavyweight Tee - Black - M', 10]);
    ws.addRow(['Ribbed Athletic Crew Socks - Black', 5]);
  });
  const book = fileWorkbook(path);

  assert.deepEqual(await book.products(), [
    { sku: '', name: 'Oversized Heavyweight Tee - Black - M' },
    { sku: '', name: 'Ribbed Athletic Crew Socks - Black' },
  ]);

  // Exact, case-insensitive, against a real cell — and the model's invented SKU cannot help.
  const { rows, missing } = await book.resolveRows([
    { description: 'ribbed athletic crew socks - black' },
    { description: 'made-up name', sku: 'ACC-SOC-BLK-OS' },
  ]);
  assert.deepEqual([...rows], [[0, 3]]);
  assert.deepEqual(missing, ['made-up name']);

  // And the whole row interface still works: draft, ask, answer.
  await book.writeDraft('5 · Ribbed Athletic Crew Socks - Black', { row: 3, note: 'a mail' });
  assert.equal((await readSheet(path)).getCell('C3').value, '5 · Ribbed Athletic Crew Socks - Black');
  assert.equal(await fileWorkbook(path).readApproval(3), 'pending');
});

test('a sheet with neither a SKU nor a product-name column says so in one sentence', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['Ref', 'Qty']);
    ws.addRow(['A-1', 10]);
  });

  await assert.rejects(fileWorkbook(path).products(), /no product column/);
  await assert.rejects(fileWorkbook(path).resolveRows([{ description: 'anything' }]), /no product column/);
});

test('a name-only table is found even when a cover sheet comes first', async () => {
  const path = await buildWorkbook((wb) => {
    wb.addWorksheet('Cover').addRow(['Example GmbH', 'price list']);
    const ws = wb.addWorksheet('Stock');
    ws.addRow(['Artikelbezeichnung', 'Qty']);
    ws.addRow(['Ribbed Athletic Crew Socks - Black', 5]);
  });

  assert.deepEqual(await fileWorkbook(path).products(), [
    { sku: '', name: 'Ribbed Athletic Crew Socks - Black' },
  ]);
});

// A workbook that does not look like the sample: these are the shapes a real SME file arrives in.

test('a title row above the table does not hide the table', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Stock');
    ws.addRow(['Bestellung 2026']);
    ws.addRow([]);
    ws.addRow(['SKU', 'Product Name', 'Qty']);
    ws.addRow(['AP-TSH-BLK-M', 'Oversized Heavyweight Tee - Black - M', 10]);
    ws.addRow(['ACC-SOC-BLK-OS', 'Ribbed Athletic Crew Socks - Black', 5]);
  });
  const book = fileWorkbook(path);

  assert.deepEqual(await book.products(), [
    { sku: 'AP-TSH-BLK-M', name: 'Oversized Heavyweight Tee - Black - M' },
    { sku: 'ACC-SOC-BLK-OS', name: 'Ribbed Athletic Crew Socks - Black' },
  ]);
  assert.equal((await book.resolveRows([{ description: 'ribbed athletic crew socks - black' }])).rows.get(0), 5);

  await book.writeDraft('30', { row: 5 });
  const sheet = await readSheet(path);
  assert.equal(sheet.getCell('A1').value, 'Bestellung 2026', 'the title is untouched');
  assert.equal(sheet.getCell('D3').value, 'AI suggests', 'the column header sits with the table header');
  assert.equal(sheet.getCell('D5').value, '30');
  assert.equal(sheet.getCell('E5').dataValidation?.type, 'list');
  assert.equal(sheet.getCell('D5').value, '30');
  assert.equal((await book.readApproval(5)), 'pending');
});

test('the orders sheet need not be the first one', async () => {
  const path = await buildWorkbook((wb) => {
    wb.addWorksheet('Cover').addRow(['Example GmbH']);
    const ws = wb.addWorksheet('Stock');
    ws.addRow(['SKU', 'Product Name', 'Qty']);
    ws.addRow(['AP-TSH-BLK-M', 'Oversized Heavyweight Tee - Black - M', 10]);
  });

  assert.deepEqual(await fileWorkbook(path).products(), [
    { sku: 'AP-TSH-BLK-M', name: 'Oversized Heavyweight Tee - Black - M' },
  ]);
});

test('the header may be German, and a trailing colon or star is the same header', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Lager');
    ws.addRow(['Artikelnummer*', 'Bezeichnung:', 'Bestand']);
    ws.addRow(['AP-TSH-BLK-M', 'Oversized Heavyweight Tee - Black - M', 10]);
  });

  assert.deepEqual(await fileWorkbook(path).products(), [
    { sku: 'AP-TSH-BLK-M', name: 'Oversized Heavyweight Tee - Black - M' },
  ]);
});

test('a product name stored as a link or as rich text is still its text', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['SKU', 'Product Name', 'Qty']);
    ws.getCell('A2').value = 'AP-TSH-BLK-M';
    ws.getCell('B2').value = {
      text: 'Oversized Heavyweight Tee - Black - M',
      hyperlink: 'https://example.invalid/tee',
    };
    ws.getCell('C2').value = 10;
    ws.getCell('A3').value = 'ACC-SOC-BLK-OS';
    ws.getCell('B3').value = { richText: [{ text: 'Ribbed Athletic Crew Socks - Black' }] };
    ws.getCell('C3').value = 5;
  });
  const book = fileWorkbook(path);

  assert.deepEqual(await book.products(), [
    { sku: 'AP-TSH-BLK-M', name: 'Oversized Heavyweight Tee - Black - M' },
    { sku: 'ACC-SOC-BLK-OS', name: 'Ribbed Athletic Crew Socks - Black' },
  ]);
  assert.deepEqual(
    [...(await book.resolveRows([{ description: 'Ribbed Athletic Crew Socks - Black' }])).rows],
    [[0, 3]],
  );
});

test('the AI columns land beside the table, not beside a stray note', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['SKU', 'Product Name', 'Qty']);
    ws.addRow(['AP-TSH-BLK-M', 'Oversized Heavyweight Tee - Black - M', 10]);
    ws.getCell('ZZ1').value = 'scratch';
  });

  await fileWorkbook(path).writeDraft('30', { row: 2 });

  const sheet = await readSheet(path);
  assert.equal(sheet.getCell('D1').value, 'AI suggests');
  assert.equal(sheet.getCell('E1').value, 'Your answer: OK or no');
  assert.equal(sheet.getCell('F1').value, 'AI result');
  assert.equal(sheet.getCell('D2').value, '30');
  assert.equal(sheet.getCell('ZZ1').value, 'scratch', 'the stray cell is untouched');
});

test('a workbook that already has its own Log sheet keeps it', async () => {
  const path = await buildWorkbook((wb) => {
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['SKU', 'Product Name', 'Qty']);
    ws.addRow(['AP-TSH-BLK-M', 'Oversized Heavyweight Tee - Black - M', 10]);
    wb.addWorksheet('Log').addRow(['my own notes']);
  });
  const book = fileWorkbook(path);

  await book.appendLog([
    { when: '2026-10-30 09:00', mail: '<m@x>', did: 'row 2 · AP-TSH-BLK-M · 30', answer: 'OK', result: 'recorded ✓' },
  ]);

  const theirs = await readSheetNamed(path, 'Log');
  assert.deepEqual(rowValues(theirs!.getRow(1)).slice(1), ['my own notes'], 'their sheet is not ours to write in');

  const ours = await readSheetNamed(path, 'Log (sheetflow)');
  assert.deepEqual(rowValues(ours!.getRow(1)).slice(1), [
    'When (UTC)',
    'Mail',
    'What the AI did',
    'Your answer',
    'Result',
  ]);
  assert.deepEqual(rowValues(ours!.getRow(2)).slice(1), [
    '2026-10-30 09:00',
    '<m@x>',
    'row 2 · AP-TSH-BLK-M · 30',
    'OK',
    'recorded ✓',
  ]);

  // Idempotent: a second pass appends to the same sheet instead of creating another one.
  await book.appendLog([{ when: '2026-10-30 09:05', mail: '<n@x>', did: 'row 2', answer: 'no', result: 'cleared' }]);
  const again = await readSheetNamed(path, 'Log (sheetflow)');
  assert.equal(again!.rowCount, 3);
});

test('a cell reads as text whatever shape Excel stored it in', () => {
  assert.equal(cellTextValue('  x  '), 'x');
  assert.equal(cellTextValue(12), '12');
  assert.equal(cellTextValue(null), '');
  assert.equal(cellTextValue({ text: 'Tee', hyperlink: 'https://example.invalid' }), 'Tee');
  assert.equal(cellTextValue({ richText: [{ text: 'Ri' }, { text: 'bbed' }] }), 'Ribbed');
  assert.equal(cellTextValue({ formula: 'A1&B1', result: 'Tee' }), 'Tee');
  assert.equal(cellTextValue({ formula: 'A1&B1' }), '', 'a formula with no cached result has no text');
});

test('a locked workbook is recognised as such, and nothing else is', () => {
  assert.equal(isSheetBusy(Object.assign(new Error('x'), { code: 'EBUSY' })), true);
  assert.equal(isSheetBusy(Object.assign(new Error('x'), { code: 'EPERM' })), true);
  assert.equal(isSheetBusy(new Error('EBUSY: resource busy or locked, open orders.xlsx')), true);
  assert.equal(
    isSheetBusy(new Error('The process cannot access the file because it is being used by another process')),
    true,
  );
  assert.equal(isSheetBusy(new Error('Unexpected end of zip file')), false);
  assert.equal(isSheetBusy('nope'), false);
});

test('parseDecision accepts what a human would actually type', () => {
  for (const yes of ['OK', 'ok.', 'yes', 'Approve', 'y', '1', '✓']) {
    assert.equal(parseDecision(yes), 'approved', yes);
  }
  for (const no of ['NO', 'no.', 'reject', 'x', '0', '✗']) {
    assert.equal(parseDecision(no), 'rejected', no);
  }
  for (const maybe of ['', '  ', '?', 'check with Sam']) {
    assert.equal(parseDecision(maybe), 'pending', maybe);
  }
});
