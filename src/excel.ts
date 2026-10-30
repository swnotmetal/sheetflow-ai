// The only thing sheetflow touches is the user's own workbook — and it is also the only interface
// the human ever sees. No terminal, no web page, no email client. So the columns have to explain
// themselves: `AI suggests`, `Your answer: OK or no`, `AI result`.
//
// The guarantees, implemented over `exceljs`:
//   - atomic writes (temp file -> rename) with a `.bak`, and `revertLastWrite`
//   - formulas and number formats survive, because existing cells are never rewritten
//   - the AI writes only into its own three columns; business columns keep their human data

// A real workbook rarely has a clean header row: a title above the table, a German or Chinese label,
// a trailing colon. `src/grid.ts` holds those rules once, so the local client and the Google Sheets
// client recognise the same header in the same way.

import ExcelJS from 'exceljs';
import { copyFile, rename } from 'node:fs/promises';
import {
  cellTextValue,
  findDataSheet,
  headerBlockEnd as headerBlockEndOf,
  headerColumn as headerColumnOf,
  keyColumnError,
  NAME_HEADERS,
  normaliseHeader,
  parseDecision,
  SKU_HEADERS,
  WorkbookShapeError,
  type GridView,
} from './grid.js';

export interface DraftCell {
  row: number;
  /**
   * Provenance written as a cell note on the suggestion: which mail asked for this. A shop owner
   * with the mail open cannot align "30 pcs" with anything otherwise. Empty means "no note".
   */
  note?: string;
}

/** What the model thinks one requested item is. */
export interface LineRef {
  description: string;
  /** A SKU the model picked from the catalogue we showed it. Verified against the sheet. */
  sku?: string | null;
}

export interface RowMatch {
  /** Index of the requested line → the sheet row it belongs to. */
  rows: Map<number, number>;
  /** Descriptions that matched nothing in the sheet. Reported, never guessed. */
  missing: string[];
}

export interface Product {
  /** The sheet's own key. Empty when the workbook has no SKU column and matches by name instead. */
  sku: string;
  name: string;
}

/** The three columns the AI owns, resolved to real column numbers for this workbook. */
type AiColumns = { draft: number; review: number; result: number };

/** One line of the audit the human can read: what the AI did, what they answered, what happened. */
export interface LogEntry {
  /** UTC, `YYYY-MM-DD HH:mm`. */
  when: string;
  mail: string;
  did: string;
  answer: string;
  result: string;
}

export interface Workbook {
  /**
   * Create the three AI columns and the `Log` sheet if they are missing. Called once per pass, so
   * the interface is visible as soon as the tool runs — waiting for the first mail to reveal where
   * to type is how a user ends up staring at a sheet with nothing in it. A no-op when present.
   */
  ensureInterface(): Promise<void>;
  /** The catalogue, read from the sheet itself — the model is never told about products we dropped. */
  products(): Promise<Product[]>;
  /**
   * Map each requested line to its own row: by SKU first, then by exact SKU or product name. A sheet
   * with no SKU column is matched by name alone — plenty of small shops only keep product names.
   */
  resolveRows(lines: LineRef[]): Promise<RowMatch>;
  /** Write a draft value for one row, creating the AI columns if they are absent. */
  writeDraft(value: string, cell: DraftCell): Promise<void>;
  /** Read the human's decision for a row from the answer column. */
  readApproval(row: number): Promise<'approved' | 'rejected' | 'pending'>;
  /** Clear a drafted row (used when the human rejects it). */
  clearDraft(row: number): Promise<void>;
  /** Tell the human what happened to their answer, in the row they answered on. */
  markResults(results: { row: number; text: string }[]): Promise<void>;
  /** What the result cell currently says — used to notice a receipt the workbook lost. */
  readResult(row: number): Promise<string>;
  /**
   * Append to the `Log` sheet. The working columns are transient by nature — a rejected draft is
   * cleared, a re-run overwrites — so the history has to live somewhere that is only ever appended
   * to. This is that place, and it is in the same file the human already has open.
   */
  appendLog(entries: LogEntry[]): Promise<void>;
  /** Undo the most recent write, using the `.bak` taken just before it. */
  revertLastWrite(): Promise<void>;
}

// Plain words, because the person filling this in is not the person who wrote the code.
const DRAFT_HEADERS = ['AI suggests', 'AI_Draft'];
const REVIEW_HEADERS = ['Your answer: OK or no', 'AI_Review'];
const RESULT_HEADERS = ['AI result', 'AI_Result'];

const LOG_SHEET = 'Log';
const LOG_HEADERS = ['When (UTC)', 'Mail', 'What the AI did', 'Your answer', 'Result'];

// What an unanswered suggestion looks like. Measured 2026-10-28: with a dropdown on every row the
// human answered rows 2–5 — the first products on screen — while the drafts sat on rows 9–24, so the
// sheet invited input where there was nothing to answer. Now only a row that carries a suggestion is
// highlighted and only that row offers the choice.
const ATTENTION = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } } as const;
const ANSWER_PICKER: ExcelJS.DataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"OK,no"'],
  showErrorMessage: true,
  errorTitle: 'OK or no',
  error: 'Type OK to accept the AI suggestion, or no to reject it.',
};

const BACKUP_SUFFIX = '.bak';

/**
 * Excel keeps the workbook open for writing, so replacing it fails with EBUSY while somebody has it
 * on screen. That is a normal, recoverable condition — not a bug — and callers need to tell it apart
 * from a malformed file.
 */
export function isSheetBusy(error: unknown): boolean {
  // Frameworks wrap step errors (Mastra does), so the real code sits on the cause chain.
  let current: unknown = error;
  for (let depth = 0; depth < 5 && current; depth++) {
    const record = current as { code?: string; message?: string; cause?: unknown };
    const code = record.code ?? '';
    if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') return true;
    const message = typeof record.message === 'string' ? record.message : String(record);
    if (/busy or locked|being used by another process/i.test(message)) return true;
    current = record.cause;
  }
  return false;
}

/** Normalise whatever a human types into the answer column — the rule lives in `src/grid.ts`. */
export { parseDecision };

/**
 * The text in a cell, whatever shape Excel stored it in — see `src/grid.ts`, where the rule lives
 * so that the Google Sheets client reads a cell the same way. Re-exported here because this module
 * is where callers expect it.
 */
export { cellTextValue };

/** A worksheet as the shared, spreadsheet-agnostic grid the matching rules are written against. */
function worksheetGrid(ws: ExcelJS.Worksheet): GridView {
  return {
    rows: ws.rowCount,
    cols: ws.columnCount,
    text: (row, col) => cellTextValue(ws.getCell(row, col).value),
  };
}

/** Which column carries one of these header names, on this worksheet. */
function headerColumnAt(
  ws: ExcelJS.Worksheet,
  headerRow: number,
  names: string[],
): number | undefined {
  return headerColumnOf(worksheetGrid(ws), headerRow, names);
}

/** The right edge of the table's header block on this worksheet. */
function headerBlockEnd(ws: ExcelJS.Worksheet, headerRow: number): number {
  return headerBlockEndOf(worksheetGrid(ws), headerRow);
}

/**
 * The sheet the orders live on, and the row its header sits in. A cover sheet, a price list or a
 * summary tab in front of the stock table is normal, so the first worksheet is only the fallback.
 */
function dataSheet(wb: ExcelJS.Workbook, path: string): { ws: ExcelJS.Worksheet; headerRow: number } {
  const found = findDataSheet(
    wb.worksheets.map((ws) => ({ name: ws.name, grid: worksheetGrid(ws) })),
  );
  const ws = found ? wb.getWorksheet(found.name) : wb.worksheets[0];
  if (!ws) throw new Error(`${path} has no worksheet`);
  return { ws, headerRow: found?.headerRow ?? 1 };
}

/** Defined in `src/grid.ts`; re-exported so `inbox.ts` and the tests keep importing it from here. */
export { WorkbookShapeError };

/**
 * exceljs has no `removeNote`, and `cell.note = undefined` still leaves an empty comment box, so the
 * slot is cleared directly — the same reason `settled()` reaches into the validation model.
 */
function clearNote(cell: ExcelJS.Cell) {
  (cell as unknown as { _comment?: unknown })._comment = undefined;
}

/** `solid` is the highlight that means "this row is asking for an answer"; anything else is quiet. */
function isSolid(fill: ExcelJS.Fill | undefined): boolean {
  return (fill as { pattern?: string } | undefined)?.pattern === 'solid';
}

/** A sheet is ours only when its first row is our header — a user's own `Log` sheet is not. */
function isLogSheet(ws: ExcelJS.Worksheet): boolean {
  if (ws.rowCount === 0) return false;
  return LOG_HEADERS.every(
    (header, at) => normaliseHeader(ws.getCell(1, at + 1).value) === normaliseHeader(header),
  );}

/**
 * The history sheet: the one whose header row is ours, else a new one under a free name. Measured
 * 2026-10-30: a workbook that already had a sheet called `Log` received the audit rows in its own
 * table, because the name was looked up instead of the header.
 */
function logSheetOf(wb: ExcelJS.Workbook): ExcelJS.Worksheet {
  const ours = wb.worksheets.find(isLogSheet);
  if (ours) return ours;
  // An empty sheet we created before crashing is still ours, and needs its header.
  const empty = wb.getWorksheet(LOG_SHEET);
  if (empty && empty.rowCount === 0) return empty;
  const taken = new Set(wb.worksheets.map((ws) => ws.name.toLowerCase()));
  let name = LOG_SHEET;
  for (let next = 2; taken.has(name.toLowerCase()); next++) {
    name = next === 2 ? `${LOG_SHEET} (sheetflow)` : `${LOG_SHEET} (sheetflow ${next - 1})`;
  }
  return wb.addWorksheet(name);
}

/**
 * A workbook backed by a real `.xlsx` file. Nothing is written until a write method is called, and
 * every flush copies the previous file to `<name>.bak` first.
 */
export function fileWorkbook(path: string): Workbook {
  // No cache on purpose. While the run is suspended the human edits this same file in Excel, so
  // every read has to see the file as it is now, not as it was when the process started.
  const sheet = async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    return { wb, ...dataSheet(wb, path) };
  };

  /** The AI's three columns, created on first use and found by either their new or old header. */
  const aiColumns = (ws: ExcelJS.Worksheet, headerRow: number) => {
    const find = (names: string[], label: string) => {
      const existing = headerColumnAt(ws, headerRow, names);
      if (existing) return existing;
      const col = headerBlockEnd(ws, headerRow) + 1;
      ws.getCell(headerRow, col).value = label;
      return col;
    };
    return {
      draft: find(DRAFT_HEADERS, DRAFT_HEADERS[0]),
      review: find(REVIEW_HEADERS, REVIEW_HEADERS[0]),
      result: find(RESULT_HEADERS, RESULT_HEADERS[0]),
    };
  };

  const flush = async (wb: ExcelJS.Workbook) => {
    // Edited cells invalidate the results Excel cached; ask Excel to recalculate on open instead of
    // leaving a stale Current Stock / Status behind.
    wb.calcProperties.fullCalcOnLoad = true;
    const tmp = `${path}.tmp`;
    await wb.xlsx.writeFile(tmp);
    await copyFile(path, path + BACKUP_SUFFIX);
    // Windows keeps the handle for a moment after Excel closes, so a replace that fails on the first
    // try is usually not a real lock. Retry briefly before telling the caller the file is busy.
    for (let attempt = 1; ; attempt++) {
      try {
        await rename(tmp, path); // replaces on both Windows and POSIX
        return;
      } catch (error) {
        if (!isSheetBusy(error) || attempt >= 5) throw error;
        await new Promise<void>((done) => setTimeout(done, 300));
      }
    }
  };

  const cellText = (ws: ExcelJS.Worksheet, row: number, col: number) =>
    cellTextValue(ws.getCell(row, col).value);

  /** A row that has a suggestion is the only row that asks for an answer. */
  const answerable = (ws: ExcelJS.Worksheet, cols: Pick<AiColumns, 'draft' | 'review'>, row: number) => {
    ws.getCell(row, cols.draft).font = { bold: true };
    ws.getCell(row, cols.draft).fill = ATTENTION;
    const answer = ws.getCell(row, cols.review);
    answer.fill = ATTENTION;
    answer.dataValidation = ANSWER_PICKER;
  };

  /** Answered or rejected: the row stops asking, and its highlight goes with it. */
  const settled = (ws: ExcelJS.Worksheet, cols: Pick<AiColumns, 'draft' | 'review'>, row: number) => {
    for (const col of [cols.draft, cols.review]) {
      ws.getCell(row, col).fill = { type: 'pattern', pattern: 'none' };
      ws.getCell(row, col).font = {};
    }
    // exceljs keeps validations in a sheet-level map keyed by address, and its `remove` only blanks
    // the value — which still serialises an empty rule. Deleting the key is what actually takes the
    // picker off the cell.
    const validations = (ws as unknown as { dataValidations?: { model?: Record<string, unknown> } })
      .dataValidations;
    if (validations?.model) delete validations.model[ws.getCell(row, cols.review).address];
  };

  return {
    async ensureInterface() {
      const { wb, ws, headerRow } = await sheet();
      const columns = ws.columnCount;
      const hadLog = wb.worksheets.some(isLogSheet);
      const { draft, review, result } = aiColumns(ws, headerRow);
      let changed = ws.columnCount !== columns || !hadLog;

      // The three columns have to look like something a person is meant to use, not like spare
      // spreadsheet space: a dark header, and — only where a suggestion exists — a highlighted cell
      // with a dropdown so nobody has to guess what to type.
      for (const col of [draft, review, result]) {
        const head = ws.getCell(headerRow, col);
        if (!head.font?.bold) {
          head.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          head.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
          changed = true;
        }
      }
      // Backfill: a suggestion written before this highlight existed still needs to be answerable.
      // The reverse matters too — a row with no suggestion must not offer a choice.
      for (let row = headerRow + 1; row <= ws.rowCount; row++) {
        const suggestion = cellText(ws, row, draft);
        const asks = suggestion !== '';
        const picker = ws.getCell(row, review).dataValidation;
        const lit = isSolid(ws.getCell(row, draft).fill);
        if (asks && (!picker || !lit)) {
          answerable(ws, { draft, review }, row);
          changed = true;
        } else if (!asks && (picker || lit)) {
          settled(ws, { draft, review }, row);
          changed = true;
        }
      }

      const log = logSheetOf(wb);
      if (log.rowCount === 0) log.addRow(LOG_HEADERS);
      if (changed) await flush(wb);
    },

    async products(): Promise<Product[]> {
      const { ws, headerRow } = await sheet();
      const skuCol = headerColumnAt(ws, headerRow, SKU_HEADERS);
      const nameCol = headerColumnAt(ws, headerRow, NAME_HEADERS);
      if (!skuCol && !nameCol) throw new WorkbookShapeError(keyColumnError(path));
      const products: Product[] = [];
      for (let row = headerRow + 1; row <= ws.rowCount; row++) {
        // A SKU column still decides which rows are products; without one, the name does.
        const key = skuCol ? cellText(ws, row, skuCol) : cellText(ws, row, nameCol!);
        if (key) products.push({ sku: skuCol ? key : '', name: nameCol ? cellText(ws, row, nameCol) : '' });
      }
      return products;
    },

    async resolveRows(lines: LineRef[]): Promise<RowMatch> {
      const { ws, headerRow } = await sheet();
      const skuCol = headerColumnAt(ws, headerRow, SKU_HEADERS);
      const nameCol = headerColumnAt(ws, headerRow, NAME_HEADERS);
      if (!skuCol && !nameCol) throw new WorkbookShapeError(keyColumnError(path));

      // Every lookup is an exact, case-insensitive match against a real cell. The model may propose
      // a SKU; if that SKU is not in this sheet, the line is unmatched — a proposal is not a fact.
      // With no SKU column there is nothing to propose against, so the product name is the key.
      const index = new Map<string, number>();
      const nameIndex = new Map<string, number>();
      for (let row = headerRow + 1; row <= ws.rowCount; row++) {
        if (skuCol) {
          const sku = cellText(ws, row, skuCol);
          if (sku && !index.has(sku.toUpperCase())) index.set(sku.toUpperCase(), row);
        }
        if (nameCol) {
          const name = cellText(ws, row, nameCol);
          if (name && !nameIndex.has(name.toUpperCase())) nameIndex.set(name.toUpperCase(), row);
        }
      }

      const rows = new Map<number, number>();
      const missing: string[] = [];
      lines.forEach((line, at) => {
        const sku = line.sku?.trim().toUpperCase();
        const description = line.description.trim().toUpperCase();
        const row =
          (sku ? index.get(sku) : undefined) ??
          index.get(description) ??
          nameIndex.get(description);
        if (row) rows.set(at, row);
        else missing.push(line.description);
      });
      return { rows, missing };
    },

    async writeDraft(value: string, cell: DraftCell) {
      const { wb, ws, headerRow } = await sheet();
      const cols = aiColumns(ws, headerRow);
      const suggestion = ws.getCell(cell.row, cols.draft);
      suggestion.value = value;
      // A note from an older mail must not survive a re-draft: it would name the wrong sender.
      if (cell.note?.trim()) suggestion.note = cell.note;
      else clearNote(suggestion);
      // A new suggestion deserves a new decision: an `OK` typed against the previous number is not
      // an `OK` for this one. (Measured 2026-10-28: answers left on rows with no suggestion would
      // otherwise be read back as consent the moment such a row was drafted again.)
      ws.getCell(cell.row, cols.review).value = null;
      answerable(ws, cols, cell.row);
      await flush(wb);
    },

    async clearDraft(row: number) {
      const { wb, ws, headerRow } = await sheet();
      const cols = aiColumns(ws, headerRow);
      ws.getCell(row, cols.draft).value = null;
      clearNote(ws.getCell(row, cols.draft));
      settled(ws, cols, row);
      await flush(wb);
    },

    async readApproval(row: number) {
      const { ws, headerRow } = await sheet();
      const review = headerColumnAt(ws, headerRow, REVIEW_HEADERS);
      if (!review) return 'pending';
      return parseDecision(cellText(ws, row, review));
    },

    async readResult(row: number) {
      const { ws, headerRow } = await sheet();
      const result = headerColumnAt(ws, headerRow, RESULT_HEADERS);
      return result ? cellText(ws, row, result) : '';
    },

    async markResults(results) {
      if (results.length === 0) return;
      const { wb, ws, headerRow } = await sheet();
      const cols = aiColumns(ws, headerRow);
      for (const { row, text } of results) {
        ws.getCell(row, cols.result).value = text;
        settled(ws, cols, row);
      }
      await flush(wb);
    },

    async appendLog(entries) {
      if (entries.length === 0) return;
      const { wb } = await sheet();
      const log = logSheetOf(wb);
      if (log.rowCount === 0) log.addRow(LOG_HEADERS);
      for (const entry of entries) {
        log.addRow([entry.when, entry.mail, entry.did, entry.answer, entry.result]);
      }
      await flush(wb);
    },

    async revertLastWrite() {
      // The `.bak` is the state just before the last flush, so this undoes that one write.
      await copyFile(path + BACKUP_SUFFIX, path);
    },
  };
}

export const workbook: Workbook = fileWorkbook(process.env.WORKBOOK_PATH ?? './orders.xlsx');
