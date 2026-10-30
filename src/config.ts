// The file a shop owner can actually edit, in place of a wall of environment variables.
//
// `config.json` sits next to the executable and answers the three questions the tool asks: which
// workbook, which mailbox, which model. Environment variables still win, so a developer's shell (and
// every test) behaves exactly as before — the file only fills in what is not already set.
//
// Paths inside the file are relative to the file's own folder, not to wherever the process happened
// to start. A shortcut, Task Scheduler, or a double-click all begin in different directories; the
// owner's answer to "which workbook" must not depend on that.

import { existsSync, readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { dirname, isAbsolute, join, resolve } from 'node:path';
import { revealSecret } from './secret.js';

export interface SheetflowConfig {
  /** Defaults to `orders.xlsx` beside this file. */
  workbook?: string;
  /** `aws` · `openai` · `anthropic` · `groq` · `deepseek` · `local` — one per run, never implied. */
  llmProfile?: string;
  llmApiKey?: string;
  /** Local model only (`llmProfile: "local"`). */
  ollama?: { baseUrl?: string; model?: string };
  /** Read-only IMAP. Leave `host` empty to read `.eml` files from `mailDir` instead. */
  imap?: {
    host?: string;
    port?: number;
    user?: string;
    /**
     * A password written in the clear. Supported so a developer and a test can work, but the owner
     * should use `set-password`, which fills `passwordProtected` instead.
     */
    password?: string;
    /** The same password, encrypted by Windows for this user (see `src/secret.ts`). */
    passwordProtected?: string;
    mailbox?: string;
    limit?: number;
  };
  /** A folder of `.eml` files, for trying it without a mailbox. */
  mailDir?: string;
  /**
   * A Google Sheet instead of a local file. Same three columns, same contract; a shared sheet is
   * synchronised, so "save and close" and "one editor" stop being questions.
   */
  googleSheet?: {
    spreadsheetId?: string;
    /** Pin the data tab. Omitted, the tab with a product header is found. */
    sheetName?: string;
    /** The service account key file (the one whose `client_email` must be an editor of the sheet). */
    serviceAccountFile?: string;
  };
  /** How long one pass waits for an answer before leaving the rows for the next pass. */
  reviewTimeoutMs?: number;
  pollMs?: number;
}

export const CONFIG_NAME = 'config.json';

/**
 * Where `config.json` and the workbook live. Next to the executable when packaged (a single
 * executable reports its own path as `process.execPath` and has no meaningful working directory),
 * otherwise the working directory, which is what `npm run inbox` means.
 */
export function configDir(env: NodeJS.ProcessEnv = process.env): string {
  const explicit = env.SHEETFLOW_CONFIG_DIR?.trim();
  if (explicit) return resolve(explicit);
  // `process.execPath` is node itself when running from source, so only trust it when packaged.
  if ((process as { isSea?: () => boolean }).isSea?.() || env.SHEETFLOW_PACKAGED === '1') {
    return dirname(process.execPath);
  }
  return process.cwd();
}

/** An absolute path for something written in the config file. */
function resolveFrom(dir: string, value: string): string {
  return isAbsolute(value) ? value : resolve(dir, value);
}

/**
 * Read the file, if there is one. A malformed file is reported as a sentence — silently falling back
 * to defaults would point the tool at the wrong workbook, which is worse than refusing to start.
 */
export function readConfig(dir: string = configDir()): SheetflowConfig | undefined {
  const path = join(dir, CONFIG_NAME);
  if (!existsSync(path)) return undefined;
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(path, 'utf8'));
  } catch (error) {
    throw new Error(
      `${path} is not valid JSON (${error instanceof Error ? error.message : String(error)}). ` +
        'Fix it, or delete it to use the defaults.',
    );
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`${path} must contain a single { … } object.`);
  }
  return parsed as SheetflowConfig;
}

/** A config file rendered as the environment variables the rest of the code already reads. */
export function envFromConfig(config: SheetflowConfig, dir: string): Record<string, string> {
  const env: Record<string, string> = {};
  const set = (key: string, value: string | number | undefined) => {
    if (value === undefined || value === null || `${value}`.trim() === '') return;
    env[key] = `${value}`;
  };

  set('WORKBOOK_PATH', config.workbook ? resolveFrom(dir, config.workbook) : join(dir, 'orders.xlsx'));
  set('LLM_PROFILE', config.llmProfile);
  set('LLM_API_KEY', config.llmApiKey);
  set('OLLAMA_BASE_URL', config.ollama?.baseUrl);
  set('OLLAMA_MODEL', config.ollama?.model);
  set('IMAP_HOST', config.imap?.host);
  set('IMAP_PORT', config.imap?.port);
  set('IMAP_USER', config.imap?.user);
  set('IMAP_PASS', config.imap?.password);
  set('IMAP_MAILBOX', config.imap?.mailbox);
  set('IMAP_LIMIT', config.imap?.limit);
  set('MAIL_DIR', config.mailDir ? resolveFrom(dir, config.mailDir) : join(dir, 'mail'));
  set(
    'SHEETFLOW_GOOGLE_SHEET',
    config.googleSheet?.spreadsheetId
      ? [config.googleSheet.spreadsheetId, config.googleSheet.sheetName ?? ''].join('|')
      : undefined,
  );
  set(
    'SHEETFLOW_GOOGLE_KEY',
    config.googleSheet?.serviceAccountFile
      ? resolveFrom(dir, config.googleSheet.serviceAccountFile)
      : undefined,
  );
  set('SHEETFLOW_REVIEW_TIMEOUT_MS', config.reviewTimeoutMs);
  set('SHEETFLOW_POLL_MS', config.pollMs);
  // The ledger belongs beside the config file too: a Task Scheduler run starts in C:\Windows\System32.
  set('SHEETFLOW_DATA_DIR', join(dir, '.data'));
  set('SHEETFLOW_MASTRA_DB', join(dir, '.data', 'mastra.db'));
  return env;
}

/**
 * Make the file's answers visible to the rest of the program. `process.env` is the one interface the
 * pipeline already speaks — mail, model credentials and the ledger all read it — so the config file
 * is turned into env defaults here rather than threaded through every function. Values already in
 * the environment are never overwritten: a developer's shell, and a test, outrank the file.
 *
 * A password stored by `set-password` is decrypted here and nowhere else, so the plaintext exists
 * only in this process's environment for the length of one run.
 */
export async function applyConfig(dir: string = configDir()): Promise<SheetflowConfig | undefined> {
  const config = readConfig(dir);
  if (!config) return undefined;
  for (const [key, value] of Object.entries(envFromConfig(config, dir))) {
    if (!process.env[key]?.trim()) process.env[key] = value;
  }
  if (config.imap?.passwordProtected && !process.env.IMAP_PASS?.trim()) {
    process.env.IMAP_PASS = await revealSecret(config.imap.passwordProtected);
  }
  return config;
}

/**
 * Store a password the way the owner should: encrypted, in place of any plaintext one. The rest of
 * the file is preserved, because it holds their answers about the workbook and the mailbox.
 */
export async function storePassword(blob: string, dir: string = configDir()): Promise<string> {
  const path = join(dir, CONFIG_NAME);
  const config = readConfig(dir) ?? {};
  const imap = { ...(config.imap ?? {}) };
  delete imap.password;
  imap.passwordProtected = blob;
  const next: SheetflowConfig = { ...config, imap };
  await writeFile(path, `${JSON.stringify(next, null, 2)}\n`, 'utf8');
  return path;
}
