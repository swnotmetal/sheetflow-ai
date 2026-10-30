// The same pipeline, against a shared Google Sheet instead of a file on one machine.
//
// This is the **start** of that direction, and it is deliberately a thin one: it implements the same
// `Workbook` seam `src/excel.ts` implements, so the graph, the ledger, the waiting and the audit do
// not change at all. What is different is the physical file, and it removes three things the local
// direction could only work around:
//
//   - Excel holding the file open  → nothing to hold; Google owns the document
//   - "save and close, then run again" → the answer is visible the moment it is typed
//   - one editor, one machine      → a shared sheet is synchronised by construction
//
// What it does **not** do yet, on purpose (a start, not a port):
//   - no dropdown, no highlight: cell formatting needs its own batchUpdate calls
//   - no `.bak` and no `revertLastWrite`: Drive keeps revisions, so undo belongs there
//   - no local file at all: the workbook is the spreadsheet, addressed by id
//
// The seam is `SheetsTransport`: everything below is logic over "rows and columns of text", so it is
// tested against a fake without a network, a credential, or a Google project.

import { createSign } from 'node:crypto';
import {
  cellTextValue,
  findDataSheet,
  headerBlockEnd,
  headerColumn,
  keyColumnError,
  NAME_HEADERS,
  parseDecision,
  SKU_HEADERS,
  WorkbookShapeError,
  fromArray,
  type GridView,
} from './grid.js';
import type { DraftCell, LineRef, LogEntry, Product, Workbook } from './excel.js';

/** The narrowest possible view of a spreadsheet: enough for every `Workbook` method, nothing more. */
export interface SheetsTransport {
  /** Every tab, in order, with the numeric id formatting requests need. */
  sheets(): Promise<{ name: string; id: number }[]>;
  /** The used part of one tab as text; `grid[0]` is row 1. */
  grid(sheet: string): Promise<string[][]>;
  /** One cell's text. Polling for an answer should not read the whole sheet every two seconds. */
  cell(sheet: string, row: number, col: number): Promise<string>;
  write(sheet: string, row: number, col: number, text: string): Promise<void>;
  /** A cell note, which is where "which mail asked for this" lives. `null` clears it. */
  note(sheet: string, row: number, col: number, text: string | null): Promise<void>;
  append(sheet: string, values: string[]): Promise<void>;
  addSheet(name: string, columns: number): Promise<void>;
}

/** `1 → A`, `27 → AA`. A spreadsheet addresses columns by letter, and so must we. */
export function columnName(col: number): string {
  let name = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(((n - 1) % 26) + 65) + name;
  }
  return name;
}

/** A tab name in a range expression. Apostrophes are the escape, exactly as the UI shows them. */
export function quoteSheet(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

const DRAFT_HEADERS = ['AI suggests', 'AI_Draft'];
const REVIEW_HEADERS = ['Your answer: OK or no', 'AI_Review'];
const RESULT_HEADERS = ['AI result', 'AI_Result'];
const LOG_SHEET = 'Log';
const LOG_HEADERS = ['When (UTC)', 'Mail', 'What the AI did', 'Your answer', 'Result'];

export interface SheetWorkbookOptions {
  /** The spreadsheet's id, from its URL. */
  spreadsheetId: string;
  transport: SheetsTransport;
  /** Pin the data tab. Omitted, the tab with a product header is found, as in the local client. */
  sheetName?: string;
}

/** A sheet function against a Google Sheet — the same contract the `.xlsx` client honours. */
export function sheetWorkbook(options: SheetWorkbookOptions): Workbook {
  const { transport } = options;
  // No cache on purpose, exactly as in the local client: the human edits this sheet between polls.

  const tabFor = async (name: string): Promise<string[][]> => transport.grid(name);

  /** The data tab and its header row, discovered the same way the local client discovers them. */
  const dataTab = async (): Promise<{ name: string; headerRow: number; grid: GridView }> => {
    if (options.sheetName) {
      const grid = fromArray(await tabFor(options.sheetName));
      const found = findDataSheet([{ name: options.sheetName, grid }]);
      return { name: options.sheetName, headerRow: found?.headerRow ?? 1, grid };
    }
    const tabs = await transport.sheets();
    const candidates: { name: string; grid: GridView }[] = [];
    for (const tab of tabs) {
      // Read one at a time: the first tab that looks like a stock table is the answer, and a cover
      // sheet in front of it is normal.
      const grid = fromArray(await tabFor(tab.name));
      candidates.push({ name: tab.name, grid });
      if (findDataSheet([{ name: tab.name, grid }])) break;
    }
    const found = findDataSheet(candidates);
    if (found) {
      const grid = candidates.find((candidate) => candidate.name === found.name)!.grid;
      return { name: found.name, headerRow: found.headerRow, grid };
    }
    const first = candidates[0] ?? { name: tabs[0]?.name ?? 'Sheet1', grid: fromArray([]) };
    return { name: first.name, headerRow: 1, grid: first.grid };
  };

  /** The three AI columns, created in the header row if they are missing. */
  const aiColumns = async (
    tab: string,
    headerRow: number,
    grid: GridView,
  ): Promise<{ draft: number; review: number; result: number }> => {
    let end = headerBlockEnd(grid, headerRow);
    const find = async (names: string[], label: string) => {
      const existing = headerColumn(grid, headerRow, names);
      if (existing) return existing;
      // Each new column widens the block, so the next one lands beside it rather than on top of it.
      end += 1;
      await transport.write(tab, headerRow, end, label);
      return end;
    };
    return {
      draft: await find(DRAFT_HEADERS, DRAFT_HEADERS[0]),
      review: await find(REVIEW_HEADERS, REVIEW_HEADERS[0]),
      result: await find(RESULT_HEADERS, RESULT_HEADERS[0]),
    };
  };

  /** Our `Log` tab is the one whose first row is our header — not a tab the user named `Log`. */
  const logTab = async (): Promise<string> => {
    for (const tab of await transport.sheets()) {
      const grid = fromArray(await tabFor(tab.name));
      if (grid.rows === 0) continue;
      const ours = LOG_HEADERS.every(
        (header, at) => cellTextValue(grid.text(1, at + 1)).toLowerCase() === header.toLowerCase(),
      );
      if (ours) return tab.name;
    }
    const taken = new Set((await transport.sheets()).map((tab) => tab.name.toLowerCase()));
    const existing = taken.has(LOG_SHEET.toLowerCase())
      ? (await transport.sheets()).find((tab) => tab.name.toLowerCase() === LOG_SHEET.toLowerCase())
      : undefined;
    if (existing) {
      const grid = fromArray(await tabFor(existing.name));
      // An empty tab of that name is one we made and never finished.
      if (grid.rows === 0) {
        await transport.append(existing.name, LOG_HEADERS);
        return existing.name;
      }
    }
    let name = LOG_SHEET;
    for (let next = 2; taken.has(name.toLowerCase()); next++) {
      name = next === 2 ? `${LOG_SHEET} (sheetflow)` : `${LOG_SHEET} (sheetflow ${next - 1})`;
    }
    await transport.addSheet(name, LOG_HEADERS.length);
    await transport.append(name, LOG_HEADERS);
    return name;
  };

  return {
    async ensureInterface() {
      const { name, headerRow, grid } = await dataTab();
      const skuCol = headerColumn(grid, headerRow, SKU_HEADERS);
      const nameCol = headerColumn(grid, headerRow, NAME_HEADERS);
      if (!skuCol && !nameCol) throw new WorkbookShapeError(keyColumnError(options.spreadsheetId));
      await aiColumns(name, headerRow, grid);
      await logTab();
    },

    async products(): Promise<Product[]> {
      const { name, headerRow, grid } = await dataTab();
      const skuCol = headerColumn(grid, headerRow, SKU_HEADERS);
      const nameCol = headerColumn(grid, headerRow, NAME_HEADERS);
      if (!skuCol && !nameCol) throw new WorkbookShapeError(keyColumnError(options.spreadsheetId));
      const products: Product[] = [];
      for (let row = headerRow + 1; row <= grid.rows; row++) {
        const key = skuCol ? grid.text(row, skuCol) : grid.text(row, nameCol!);
        if (key) products.push({ sku: skuCol ? key : '', name: nameCol ? grid.text(row, nameCol) : '' });
      }
      return products;
    },

    async resolveRows(lines: LineRef[]) {
      const { headerRow, grid } = await dataTab();
      const skuCol = headerColumn(grid, headerRow, SKU_HEADERS);
      const nameCol = headerColumn(grid, headerRow, NAME_HEADERS);
      if (!skuCol && !nameCol) throw new WorkbookShapeError(keyColumnError(options.spreadsheetId));

      // Exact, case-insensitive match against a real cell. A SKU the model invented is not a fact,
      // and a near-match is reported rather than guessed — the same rule as the local client.
      const index = new Map<string, number>();
      const nameIndex = new Map<string, number>();
      for (let row = headerRow + 1; row <= grid.rows; row++) {
        if (skuCol) {
          const sku = grid.text(row, skuCol).toUpperCase();
          if (sku && !index.has(sku)) index.set(sku, row);
        }
        if (nameCol) {
          const value = grid.text(row, nameCol).toUpperCase();
          if (value && !nameIndex.has(value)) nameIndex.set(value, row);
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
      const { name, headerRow, grid } = await dataTab();
      const cols = await aiColumns(name, headerRow, grid);
      await transport.write(name, cell.row, cols.draft, value);
      await transport.note(name, cell.row, cols.draft, cell.note?.trim() || null);
      // A new suggestion deserves a new decision: an `OK` against the previous number is not an
      // `OK` for this one.
      await transport.write(name, cell.row, cols.review, '');
    },

    async clearDraft(row: number) {
      const { name, headerRow, grid } = await dataTab();
      const cols = await aiColumns(name, headerRow, grid);
      await transport.write(name, row, cols.draft, '');
      await transport.note(name, row, cols.draft, null);
    },

    async readApproval(row: number) {
      const { name, headerRow, grid } = await dataTab();
      const review = headerColumn(grid, headerRow, REVIEW_HEADERS);
      if (!review) return 'pending';
      return parseDecision(await transport.cell(name, row, review));
    },

    async readResult(row: number) {
      const { name, headerRow, grid } = await dataTab();
      const result = headerColumn(grid, headerRow, RESULT_HEADERS);
      return result ? transport.cell(name, row, result) : '';
    },

    async markResults(results) {
      if (results.length === 0) return;
      const { name, headerRow, grid } = await dataTab();
      const cols = await aiColumns(name, headerRow, grid);
      for (const { row, text } of results) await transport.write(name, row, cols.result, text);
    },

    async appendLog(entries: LogEntry[]) {
      if (entries.length === 0) return;
      const tab = await logTab();
      for (const entry of entries) {
        await transport.append(tab, [entry.when, entry.mail, entry.did, entry.answer, entry.result]);
      }
    },

    async revertLastWrite() {
      // Drive keeps the revisions; a second copy beside the file is the local client's workaround for
      // not having them. Refusing is more honest than pretending to undo.
      throw new Error(
        'Google Sheets has no `.bak` to revert to — use File ▸ Version history in the browser.',
      );
    },
  };
}

// ------------------------------------------------------------------------------------------------
// The real transport: a service account, because it needs no consent screen and no redirect URI.
// ------------------------------------------------------------------------------------------------

export interface ServiceAccount {
  client_email: string;
  private_key: string;
}

export interface GoogleSheetsOptions {
  spreadsheetId: string;
  serviceAccount: ServiceAccount;
  /** Injected in tests; defaults to the global `fetch`. */
  http?: typeof fetch;
}

const TOKEN_URL = 'https://oauth2.googleapis.com/token';
const API = 'https://sheets.googleapis.com/v4/spreadsheets';
const SCOPE = 'https://www.googleapis.com/auth/spreadsheets';

function base64url(value: Buffer | string): string {
  return Buffer.from(value)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

/**
 * A signed JWT assertion, which is how a service account proves who it is without a browser. Built
 * with `node:crypto` so the Google client library — and its dependencies — never enter the picture.
 */
export function signedAssertion(account: ServiceAccount, now = Math.floor(Date.now() / 1000)): string {
  const header = base64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = base64url(
    JSON.stringify({
      iss: account.client_email,
      scope: SCOPE,
      aud: TOKEN_URL,
      iat: now,
      exp: now + 3600,
    }),
  );
  const signature = createSign('RSA-SHA256')
    .update(`${header}.${claim}`)
    .sign(account.private_key);
  return `${header}.${claim}.${base64url(signature)}`;
}

/** A `SheetsTransport` over the Sheets REST API, authenticated as a service account. */
export function googleSheetsTransport(options: GoogleSheetsOptions): SheetsTransport {
  const { spreadsheetId, serviceAccount } = options;
  const fetchImpl = options.http ?? fetch;
  let token: { value: string; expiresAt: number } | undefined;

  const accessToken = async (): Promise<string> => {
    if (token && Date.now() < token.expiresAt) return token.value;
    const body = new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: signedAssertion(serviceAccount),
    });
    const response = await fetchImpl(TOKEN_URL, {
      method: 'POST',
      headers: { 'content-type': 'application/x-www-form-urlencoded' },
      body,
    });
    if (!response.ok) {
      throw new Error(
        `Google refused the service account (${response.status}). Check that the key is valid and ` +
          `that the sheet is shared with ${serviceAccount.client_email} as an editor.`,
      );
    }
    const payload = (await response.json()) as { access_token: string; expires_in?: number };
    // Refresh a minute early: a token that expires mid-request fails for no visible reason.
    token = {
      value: payload.access_token,
      expiresAt: Date.now() + ((payload.expires_in ?? 3600) - 60) * 1000,
    };
    return token.value;
  };

  const call = async (url: string, init: RequestInit = {}): Promise<unknown> => {
    const response = await fetchImpl(url, {
      ...init,
      headers: {
        authorization: `Bearer ${await accessToken()}`,
        'content-type': 'application/json',
        ...(init.headers ?? {}),
      },
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(`Google Sheets ${response.status}: ${detail.slice(0, 300)}`);
    }
    return response.status === 204 ? undefined : response.json();
  };

  const range = (sheet: string, a1: string) => `${API}/${spreadsheetId}/values/${encodeURIComponent(`${quoteSheet(sheet)}!${a1}`)}`;
  let sheetIds: { name: string; id: number }[] | undefined;

  const refreshSheets = async () => {
    const data = (await call(`${API}/${spreadsheetId}?fields=sheets.properties`)) as {
      sheets?: { properties?: { sheetId?: number; title?: string } }[];
    };
    sheetIds = (data.sheets ?? []).map((sheet) => ({
      name: sheet.properties?.title ?? '',
      id: sheet.properties?.sheetId ?? 0,
    }));
    return sheetIds;
  };

  return {
    async sheets() {
      return sheetIds ?? refreshSheets();
    },

    async grid(sheet) {
      const data = (await call(
        `${range(sheet, 'A1:ZZ')}?majorDimension=ROWS&valueRenderOption=FORMATTED_VALUE`,
      )) as { values?: string[][] };
      return data.values ?? [];
    },

    async cell(sheet, row, col) {
      const data = (await call(
        `${range(sheet, `${columnName(col)}${row}`)}?valueRenderOption=FORMATTED_VALUE`,
      )) as { values?: string[][] };
      return data.values?.[0]?.[0] ?? '';
    },

    async write(sheet, row, col, text) {
      await call(`${range(sheet, `${columnName(col)}${row}`)}?valueInputOption=RAW`, {
        method: 'PUT',
        body: JSON.stringify({ values: [[text]] }),
      });
    },

    async note(sheet, row, col, text) {
      const ids = await (sheetIds ? Promise.resolve(sheetIds) : refreshSheets());
      const id = ids.find((entry) => entry.name === sheet)?.id;
      if (id === undefined) throw new Error(`no tab named "${sheet}" in this spreadsheet`);
      await call(`${API}/${spreadsheetId}:batchUpdate`, {
        method: 'POST',
        body: JSON.stringify({
          requests: [
            {
              updateCells: {
                range: {
                  sheetId: id,
                  startRowIndex: row - 1,
                  endRowIndex: row,
                  startColumnIndex: col - 1,
                  endColumnIndex: col,
                },
                fields: 'note',
                rows: [{ values: [{ note: text ?? '' }] }],
              },
            },
          ],
        }),
      });
    },

    async append(sheet, values) {
      await call(
        `${range(sheet, 'A1')}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
        { method: 'POST', body: JSON.stringify({ values: [values] }) },
      );
    },

    async addSheet(name, columns) {
      await call(`${API}/${spreadsheetId}:batchUpdate`, {
        method: 'POST',
        body: JSON.stringify({
          requests: [{ addSheet: { properties: { title: name, gridProperties: { columnCount: columns } } } }],
        }),
      });
      sheetIds = undefined; // the tab list changed
    },
  };
}
