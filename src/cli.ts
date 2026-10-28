// Minimal entrypoint.
//
//   npm run dev           → print the resolved profile (no network, no side effects)
//   npm run dev -- inbox  → one pass of the mailbox: poll → dedupe → draft → wait → audit
//
// The inbox pass is the only command that touches a mailbox or a workbook.

import { existsSync } from 'node:fs';
import { resolveProfile } from './model.js';
import { configuredMailbox } from './mail.js';
import { fileWorkbook } from './excel.js';
import { openDb } from './db.js';
import { durableMastra } from './state.js';
import { buildWorkflow } from './workflow.js';
import { pollInbox, forget, forgetWaiting } from './inbox.js';

async function main(argv: string[]) {
  const provider = resolveProfile(process.env.LLM_PROFILE);
  console.log(
    `sheetflow: profile=${provider.profile} model=${provider.model} fallbacks=[${provider.fallbacks.join(', ') || 'none'}]`,
  );

  if (argv[0] !== 'inbox') return;

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
          : `sheetflow: forgot ${forgotten.length} run(s) [${forgotten.join(', ')}] — those mails can be drafted again`,
      );
    } finally {
      ledger.close();
    }
    return;
  }

  const workbookPath = process.env.WORKBOOK_PATH ?? './orders.xlsx';
  if (!existsSync(workbookPath)) {
    // The single most likely first-run mistake: no workbook, so nothing can be drafted into.
    throw new Error(
      `WORKBOOK_PATH points at ${workbookPath}, which does not exist.\n` +
        'Start from a copy of the sample:  Copy-Item test/inventory.xlsx orders.xlsx',
    );
  }

  const db = openDb();
  const book = fileWorkbook(workbookPath);
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
        `sheetflow: ${counts.pending} mail(s) waiting on you — open ${workbookPath}, ` +
          `pick OK or no in the "Your answer: OK or no" column (row ${rows.join(', ')}), save and close, run again`,
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
try {
  await main(process.argv.slice(2));
} catch (error) {
  console.error(`sheetflow: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
}

