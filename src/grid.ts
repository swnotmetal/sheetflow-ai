// Reading a sheet as a grid, independent of which spreadsheet it is.
//
// The local `.xlsx` client and the Google Sheets client must agree on one thing above all: how a
// header is recognised, and how a cell's text is read. If those drift, a shop's file behaves one way
// on disk and another in the cloud for no visible reason. So the rules live here, once, over a
// minimal "rows × columns of text" view — no ExcelJS, no HTTP, nothing to mock.

/** A sheet seen as text, 1-based like every spreadsheet formula a human will use. */
export interface GridView {
  /** Last row that has anything in it. */
  rows: number;
  /** Last column that has anything in it. */
  cols: number;
  /** The text of one cell; `''` outside the sheet. */
  text(row: number, col: number): string;
}

/** A spreadsheet rarely looks like the sample: a title row above the table, a German or Chinese
 * header, a trailing colon. These are the names sheetflow recognises. The first column whose header
 * matches wins, so "Product Name" beats a later "Description". */
export const SKU_HEADERS = [
  'sku',
  'sku nr',
  'sku no',
  'sku#',
  'article',
  'article no',
  'article nr',
  'article number',
  'artikel',
  'artikel-nr',
  'artikelnummer',
  'item',
  'item no',
  'item code',
  'part no',
  'part number',
  'product code',
  '货号',
  '产品编号',
];
export const NAME_HEADERS = [
  'product name',
  'product',
  'product description',
  'description',
  'item name',
  'artikelbezeichnung',
  'bezeichnung',
  'produkt',
  '产品名称',
  '品名',
  '名称',
];

/** How far down the header may sit. Past this it is not a table header, it is text. */
export const HEADER_SCAN_ROWS = 20;

/**
 * The text in a cell, whatever shape the client handed over: a plain value, a formula's cached
 * result, a hyperlink (`{text, hyperlink}`) or rich text. `String(value)` on the last two gives
 * "[object Object]" — and a catalogue whose product names link to a product page is ordinary.
 */
export function cellTextValue(value: unknown): string {
  if (value == null) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value instanceof Date) return value.toISOString();
  // Rich text is an array of runs, each of which is one of the object shapes below.
  if (Array.isArray(value)) return value.map(cellTextValue).join('').trim();
  const record = value as { text?: unknown; richText?: unknown; result?: unknown; formula?: unknown };
  if (record.richText !== undefined) return cellTextValue(record.richText);
  if (record.text !== undefined) return cellTextValue(record.text);
  if (record.result !== undefined) return cellTextValue(record.result);
  // A formula whose cached result was never computed has no text to read.
  return record.formula === undefined ? String(value).trim() : '';
}

/** Headers arrive with colons, stars and doubled spaces; compare them without any of that. */
export function normaliseHeader(value: unknown): string {
  return cellTextValue(value).toLowerCase().replace(/\s+/g, ' ').replace(/[:*]+$/, '').trim();
}

export function headerColumn(grid: GridView, headerRow: number, names: string[]): number | undefined {
  const wanted = names.map(normaliseHeader);
  for (let col = 1; col <= grid.cols; col++) {
    if (wanted.includes(normaliseHeader(grid.text(headerRow, col)))) return col;
  }
  return undefined;
}

/**
 * The table header is wherever the key column lives in the first rows — usually row 1, often row 3
 * under a title. Scanning stops at `HEADER_SCAN_ROWS` so a stray word far down is never mistaken
 * for one.
 */
export function findHeaderRow(grid: GridView, names: string[]): number | undefined {
  const last = Math.min(grid.rows, HEADER_SCAN_ROWS);
  for (let row = 1; row <= last; row++) {
    if (headerColumn(grid, row, names) !== undefined) return row;
  }
  return undefined;
}

/**
 * The right edge of the table's header block. The AI's columns go right after it rather than after
 * the last used column: one stray note in `ZZ1` makes the sheet 702 columns wide and would bury them
 * off screen.
 */
export function headerBlockEnd(grid: GridView, headerRow: number): number {
  let start = 1;
  while (start <= grid.cols && grid.text(headerRow, start) === '') start++;
  if (start > grid.cols) return 0;
  let end = start;
  while (end <= grid.cols && grid.text(headerRow, end) !== '') end++;
  return end - 1;
}

/**
 * The sheet the orders live on, and the row its header sits in. A SKU column is the strongest
 * evidence, so a sheet that has one wins; only then is a product-name column enough — a cover sheet
 * is unlikely to say `SKU`, but it may well say `Description`.
 */
export function findDataSheet<N extends string>(
  sheets: { name: N; grid: GridView }[],
): { name: N; headerRow: number } | undefined {
  for (const names of [SKU_HEADERS, NAME_HEADERS]) {
    for (const sheet of sheets) {
      const headerRow = findHeaderRow(sheet.grid, names);
      if (headerRow !== undefined) return { name: sheet.name, headerRow };
    }
  }
  return undefined;
}

/**
 * The workbook itself cannot be read as a stock table — nothing about a single mail is wrong. Every
 * mail will fail until the sheet is fixed, so the driver reports it once, as a workbook problem.
 */
export class WorkbookShapeError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'WorkbookShapeError';
  }
}

/** Neither key column exists, so there is nothing to match a mail against — one actionable sentence. */
export function keyColumnError(subject: string): string {
  return (
    `${subject} has no product column.\n` +
    'Add a header called "SKU" (or "Artikelnummer", "Item No") — or, if you do not use SKUs, ' +
    'a "Product Name" column that matches how customers write the item.'
  );
}

/** Normalise whatever a human types into the answer column. Shared by both spreadsheet backends. */
export function parseDecision(raw: unknown): 'approved' | 'rejected' | 'pending' {
  const text = String(raw ?? '')
    .trim()
    .toLowerCase()
    .replace(/[.\s]+$/, '');
  if (!text) return 'pending';
  if (/^(ok|okay|yes|y|approve[d]?|true|1|✓|☑)$/.test(text)) return 'approved';
  if (/^(no|n|reject(ed)?|false|0|x|✗|☒)$/.test(text)) return 'rejected';
  return 'pending';
}

/** A grid held in memory, which is what a test and an HTTP response both are. */
export function fromArray(rows: string[][]): GridView {
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  return {
    rows: rows.length,
    cols: width,
    text(row, col) {
      if (row < 1 || col < 1) return '';
      return rows[row - 1]?.[col - 1]?.trim() ?? '';
    },
  };
}
