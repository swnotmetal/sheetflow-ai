// Minimal entrypoint.
//
//   npm run dev           → print the resolved profile (no network, no side effects)
//   npm run dev -- inbox  → one pass of the mailbox: poll → dedupe → draft → wait → audit
//
// The inbox pass is the only command that touches a mailbox or a workbook.

import { existsSync, readFileSync } from 'node:fs';
import { resolveProfile } from './model.js';
import { configuredMailbox } from './mail.js';
import { fileWorkbook, type Workbook } from './excel.js';
import { googleSheetsTransport, sheetWorkbook, type ServiceAccount } from './sheets.js';
import { openDb } from './db.js';
import { durableMastra } from './state.js';
import { buildWorkflow } from './workflow.js';
import { pollInbox, forget, forgetWaiting } from './inbox.js';
import { applyConfig, storePassword } from './config.js';
import { promptForSecret } from './secret.js';

/**
 * The workbook the pipeline drafts into: the shop's own `.xlsx` by default, or a Google Sheet when
 * `config.json` names one. Everything above this line is unaware of the difference — that is the
 * whole point of the `Workbook` seam.
 */
async function openWorkbook(): Promise<{ book: Workbook; where: string }> {
  const google = process.env.SHEETFLOW_GOOGLE_SHEET?.trim();
  if (google) {
    const [spreadsheetId, sheetName] = google.split('|');
    if (!spreadsheetId) throw new Error('config.json names a Google Sheet but no spreadsheetId.');
    const keyFile = process.env.SHEETFLOW_GOOGLE_KEY?.trim();
    if (!keyFile || !existsSync(keyFile)) {
      throw new Error(
        'config.json points at a Google Sheet but the service account key file is missing.\n' +
          `Expected it at ${keyFile ?? '(no serviceAccountFile set)'}. Download it from the Google ` +
          'Cloud console and share the sheet with that account as an editor.',
      );
    }
    const serviceAccount = JSON.parse(readFileSync(keyFile, 'utf8')) as ServiceAccount;
    return {
      book: sheetWorkbook({
        spreadsheetId,
        sheetName: sheetName || undefined,
        transport: googleSheetsTransport({ spreadsheetId, serviceAccount }),
      }),
      where: `the sheet https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
    };
  }

  const workbookPath = process.env.WORKBOOK_PATH ?? './orders.xlsx';
  if (!existsSync(workbookPath)) {
    // The single most likely first-run mistake: no workbook, so nothing can be drafted into.
    throw new Error(
      `There is no workbook at ${workbookPath}.\n` +
        `Put the shop's .xlsx there, or name the file in config.json ("workbook").`,
    );
  }
  return { book: fileWorkbook(workbookPath), where: workbookPath };
}

async function main(argv: string[]) {
  // A config file beside the executable is how a shop owner sets this up; the environment still
  // wins, so a developer's shell and every test behave exactly as before.
  await applyConfig();

  // `set-password` is the only command that may run without a working configuration: its job is to
  // fix the one thing the owner should never type into a file.
  if (argv[0] === 'set-password') {
    const blob = await promptForSecret();
    const path = await storePassword(blob);
    console.log(
      `sheetflow: the mailbox password is now encrypted for this Windows user in ${path}\n` +
        'It cannot be read on another machine or by another account.',
    );
    return;
  }

  const provider = resolveProfile(process.env.LLM_PROFILE);
  console.log(
    `sheetflow: profile=${provider.profile} model=${provider.model} fallbacks=[${provider.fallbacks.join(', ') || 'none'}]`,
  );

  if (argv[0] !== 'inbox') return;

  // One pass, no waiting. The scheduled task runs the same command every few minutes, so a pass that
  // sat waiting for a human would simply collide with the next one; the answer is picked up later.
  if (argv.includes('--no-wait')) process.env.SHEETFLOW_REVIEW_TIMEOUT_MS = '0';

  // A ledger row that never got an answer, or died on a lock, can never finish. Forget it on
  // request — it needs no workbook, only the ledger.
  if (argv[1] === 'forget-waiting' || argv[1] === 'retry-failed') {
    const ledger = openDb();
    try {
      const forgotten =
        argv[1] === 'forget-waiting'
          ? await forgetWaiting(ledger)
          : await forget(ledger, ['failed'], 'the run failed, so the mail may be drafted again');
      console.log(
        forgotten.length === 0
          ? 'sheetflow: nothing to forget'
          : `sheetflow: forgot ${forgotten.length} run(s) [${forgotten.join(', ')}] - those mails can be drafted again`,
      );
    } finally {
      ledger.close();
    }
    return;
  }

  const db = openDb();
  const { book, where } = await openWorkbook();
  const mastra = durableMastra();
  const workflow = buildWorkflow({ book, provider, mastra });

  try {
    const outcomes = await pollInbox({
      db,
      book,
      workflow,
      mailbox: configuredMailbox(),
      log: (line) => console.log(`sheetflow: inbox: ${line}`),
    });

    const counts = outcomes.reduce<Record<string, number>>((acc, outcome) => {
      acc[outcome.action] = (acc[outcome.action] ?? 0) + 1;
      return acc;
    }, {});
    console.log(
      `sheetflow: inbox: ${
        Object.entries(counts)
          .map(([action, n]) => `${action}=${n}`)
          .join(' ') || 'nothing to do'
      }`,
    );
    if (counts.pending) {
      const rows = outcomes.filter((outcome) => outcome.action === 'pending').flatMap((outcome) => outcome.rows);
      console.log(
        // ASCII only: a Windows console that is not on codepage 65001 renders an em dash as mojibake.
        `sheetflow: ${counts.pending} mail(s) waiting on you - open ${where}, ` +
          `pick OK or no in the "Your answer: OK or no" column (row ${rows.join(', ')}), then save`,
      );
    }
  } finally {
    // Both native handles must go before Node tears the environment down, or the process exits
    // with an assertion from a finaliser that runs too late.
    await mastra.shutdown({ drainTimeout: 2_000 });
    db.close();
  }
}

// A first run is usually where things are misconfigured, so the failure has to read like a sentence
// rather than a stack trace. `--stack` (or NODE_OPTIONS) still gives the full trace when debugging.
//
// Deliberately not top-level await: the packaged build is bundled as CommonJS so it can be embedded
// in a single executable, and a CJS entry cannot await at the top level.
main(process.argv.slice(2)).catch((error: unknown) => {
  console.error(`sheetflow: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
});

