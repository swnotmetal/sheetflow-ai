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
import { fileWorkbook, isSheetBusy, parseDecision } from '../src/excel.js';

const SOURCE = 'test/inventory.xlsx';

async function scratchCopy() {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-'));
  const path = join(dir, 'inventory.xlsx');
  await copyFile(SOURCE, path);
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
