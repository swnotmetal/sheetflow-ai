// The driver: one pass over the mailbox, then one pass over runs that are still waiting.
//
//   poll → dedupe by Message-ID → extract → draft → wait for the human's marks → resume → audit
//
// Three things no single step can own:
//
// 1. **Dedupe.** The mailbox is read-only, so the same message returns on every poll. The
//    `workflow_runs.message_id` UNIQUE constraint is what stops "poll again" from becoming
//    "draft the same order twice".
// 2. **Waiting.** Mastra suspends the run; the driver polls the sheet until every touched row has a
//    mark, then resumes. A run that is not fully marked stays suspended — and because the snapshot
//    is a file (`src/state.ts`), it survives the process exiting.
// 3. **Recording.** One audit line per row. Nothing here decides anything; the mark in `AI_Review`
//    is the decision.

import { z } from 'zod';
import {
  audit,
  deleteRun,
  getMailLabel,
  getRun,
  listRunsByStatus,
  markNoticesProjected,
  notice,
  pendingNotices,
  setMailLabel,
  setRunStatus,
  supersedeNotices,
  upsertRun,
  type Db,
  type RunStatus,
} from './db.js';
import type { Workbook } from './excel.js';
import { isSheetBusy } from './excel.js';
import { INTENTS, type Intent } from './model.js';
import type { Mail, Mailbox } from './mail.js';
import { EMPTY_MAIL, MailIdentity, mailSource, type buildWorkflow } from './workflow.js';

type Workflow = ReturnType<typeof buildWorkflow>;
type Run = Awaited<ReturnType<Workflow['createRun']>>;
type Decisions = { approved: number[]; rejected: number[]; pending: number[]; unmatched: string[] };
/** One answer as the graph receives it. */
type Decided = { row: number; decision: 'approved' | 'rejected' };
type Action = MailOutcome['action'];
type Wait = Required<Pick<InboxOptions, 'pollMs' | 'timeoutMs' | 'sleep'>>;

// A quantity attached to a unit is the one signal a real order always leaves in the text.
const QUANTITY_HINT = /\d[\d.,]*\s*(units?|pcs|pieces?|sets?|st(ü|u)ck|stk\.?)\b/i;

/** The one notice a successful repair makes untrue; kept in one place so the two cannot drift. */
const NOT_WRITTEN_YET = 'not written yet — close Excel and it appears';

/** Why a pass could not touch the file; said the same way everywhere it is said. */
const BLOCKED_BY_EXCEL = 'could not be drafted — the workbook was open in Excel';

/**
 * A failure as the shop owner needs to read it. The technical reason belongs in the audit trail;
 * the sheet gets either the workbook's own complaint (the fix is the sheet, not the mail) or one
 * sentence saying a person has to look at the mail.
 */
function failureNotice(detail: string, mail: Mail): {
  messageId: string;
  mail: string;
  did: string;
  result: string;
} {
  const shape = detail.includes('WorkbookShapeError');
  return {
    messageId: shape ? '(the workbook)' : mail.messageId,
    mail: shape ? '(sheetflow itself)' : mailSource(mail),
    did: shape
      ? detail.split('\n')[0].replace(/^.*WorkbookShapeError:\s*/, '')
      : 'could not be read automatically — needs a person to look at it',
    result: shape ? 'fix the workbook' : 'not drafted',
  };
}

/**
 * The technical reason, for the audit trail. Mastra reports a failed step as a plain object
 * (`{message, name, cause}`), not an `Error`, and `String()` on it gives "[object Object]" — so walk
 * the chain and keep the deepest message, which is the one that names the actual problem.
 */
function reasonText(error: unknown): string {
  let current: unknown = error;
  let name = '';
  let deepest = '';
  for (let depth = 0; depth < 5 && current; depth++) {
    if (typeof current === 'string') {
      deepest = current.trim();
      break;
    }
    const record = current as { message?: unknown; name?: unknown; cause?: unknown };
    if (typeof record.message === 'string' && record.message.trim()) {
      if (typeof record.name === 'string' && record.name.trim()) name = record.name.trim();
      deepest = record.message.trim();
    }
    current = record.cause;
  }
  return [name, deepest].filter(Boolean).join(': ');
}

/** Exactly what the review step suspends with. Validated, because it crosses a step boundary. */
const SuspendedReview = z.object({
  messageId: z.string(),
  // Defaults keep a snapshot written before the mail identity existed resumable.
  mail: MailIdentity.default(EMPTY_MAIL),
  // Defaults to `order` so a snapshot written before intent existed still resumes as it did.
  intent: z.enum(INTENTS).default('order'),
  modelIntent: z.enum(INTENTS).optional(),
  rows: z.array(z.object({ row: z.number().int(), matched: z.string(), draft: z.string() })),
  unmatched: z.array(z.string()),
  unquantified: z.array(z.string()).default([]),
});

export interface InboxOptions {
  mailbox: Mailbox;
  workflow: Workflow;
  db: Db;
  book: Workbook;
  /** How often the sheet is re-read while waiting for the human. */
  pollMs?: number;
  /** How long one pass waits before leaving the run suspended for the next pass. */
  timeoutMs?: number;
  /** Retry budget for a write-back blocked by Excel holding the file open. */
  writeBack?: { retries: number; pauseMs: number };
  sleep?: (ms: number) => Promise<void>;
  log?: (line: string) => void;
}

export interface MailOutcome {
  messageId: string;
  /** `skipped` means the ledger had already seen this message; `deferred` means the sheet was locked. */
  action: 'skipped' | 'deferred' | 'failed' | 'unmatched' | 'ignored' | 'approved' | 'rejected' | 'pending';
  rows: number[];
  unmatched: string[];
}

function suspendedReview(payload: unknown): z.infer<typeof SuspendedReview> | undefined {
  const parsed = SuspendedReview.safeParse(payload);
  return parsed.success ? parsed.data : undefined;
}

function stepsOf(result: unknown): Record<string, { suspendPayload?: unknown }> | undefined {
  return (result as { steps?: Record<string, { suspendPayload?: unknown }> }).steps;
}

/** Re-read the sheet until every row has a decision. False means the deadline passed first. */
export async function waitForDecisions(
  book: Workbook,
  rows: number[],
  { pollMs, timeoutMs, sleep }: Wait,
): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    let decided = true;
    try {
      for (const row of rows) {
        if ((await book.readApproval(row)) === 'pending') {
          decided = false;
          break;
        }
      }
    } catch (error) {
      // Excel was mid-save, or on screen. Keep waiting rather than killing the run.
      if (!isSheetBusy(error)) throw error;
      decided = false;
    }
    if (decided) return true;
    if (Date.now() >= deadline) return false;
    await sleep(pollMs);
  }
}

function statusFor({ approved, rejected, pending }: Decisions): RunStatus {
  if (pending.length > 0) return 'pending';
  return approved.length > 0 ? 'approved' : 'rejected';
}

type SuspendedRow = { row: number; matched: string; draft: string };

/** What the review step suspended with, once validated. */
type Review = z.infer<typeof SuspendedReview> & { rows: SuspendedRow[] };

function stamp(): string {
  return new Date().toISOString().slice(0, 16).replace('T', ' ');
}

/**
 * Write back to a workbook somebody may still have open in Excel. Retries rather than throwing: the
 * person needs a few seconds to close the file, and their decision must not be lost because of it.
 */
async function writeBack(
  budget: { retries: number; pauseMs: number } = { retries: 10, pauseMs: 1_000 },
  write: () => Promise<void>,
): Promise<boolean> {
  for (let attempt = 0; ; attempt++) {
    try {
      await write();
      return true;
    } catch (error) {
      if (!isSheetBusy(error) || attempt >= budget.retries) return false;
      await new Promise<void>((done) => setTimeout(done, budget.pauseMs));
    }
  }
}

/**
 * Put queued failures where the shop owner will actually see them. The sheet is the only interface,
 * so a message printed to a terminal he has already closed is a message he never received — and the
 * most common failure ("close Excel so I can write") is exactly the one that cannot be written at
 * the moment it happens. It queues in the ledger and lands on the next pass that finds the file free.
 */
async function projectNotices(
  db: Db,
  book: Workbook,
  budget: { retries: number; pauseMs: number } | undefined,
  log: (line: string) => void,
): Promise<void> {
  const pending = pendingNotices(db);
  if (pending.length === 0) return;
  const written = await writeBack(budget, () =>
    book.appendLog(
      pending.map((entry) => ({
        when: entry.when_label,
        mail: entry.mail_label,
        did: entry.did,
        answer: '—',
        result: entry.result,
      })),
    ),
  );
  // Still busy: the next pass writes it. A crash between the write and the mark repeats one line,
  // which is the harmless direction for an append-only history.
  if (!written) return;
  markNoticesProjected(db, pending.map((entry) => entry.id));
  log(`wrote ${pending.length} problem(s) into the sheet`);
}

/**
 * A run that finished while Excel held the file owes the workbook a receipt. Every pass re-applies
 * any that are still missing, once, so the sheet catches up by itself the next time it is free —
 * no command for the human to remember. Measured 2026-10-28: two approvals were recorded in the
 * ledger and never appeared in the workbook, which is exactly the kind of gap that kills trust.
 */
async function repairSheet(
  db: Db,
  book: Workbook,
  budget: { retries: number; pauseMs: number } | undefined,
  log: (line: string) => void,
): Promise<void> {
  // Two sources of doubt: a write that failed at the time, and a finished run whose decision the
  // workbook no longer shows (a human saving an older copy of the file undoes writes silently).
  const failedWrites = db
    .prepare(`SELECT DISTINCT run_id AS id FROM audit_log WHERE action LIKE 'sheet-write-failed%'`)
    .all() as { id: string }[];
  const finished = db
    .prepare(`SELECT id FROM workflow_runs WHERE status IN ('approved', 'rejected')`)
    .all() as { id: string }[];
  const owing = [...new Set([...failedWrites, ...finished].map((row) => row.id))];

  for (const runId of owing) {
    // Already repaired and still showing? Then nothing to do.
    const rows = db
      .prepare(`SELECT action FROM audit_log WHERE run_id = ? AND actor = 'human'`)
      .all(runId) as { action: string }[];
    const picked = (word: 'approved' | 'rejected') =>
      rows
        .map((row) => new RegExp(`^${word} row (\\d+)$`).exec(row.action)?.[1])
        .filter((value): value is string => value !== undefined)
        .map(Number);
    const decided = [...picked('approved'), ...picked('rejected')];
    if (decided.length === 0) continue;

    let missing: number[];
    try {
      const seen = await Promise.all(
        decided.map(async (row) => ((await book.readResult(row)).trim() ? null : row)),
      );
      missing = seen.filter((row): row is number => row !== null);
    } catch (error) {
      if (isSheetBusy(error)) continue; // try again next pass
      throw error;
    }
    if (missing.length === 0) continue;

    const approved = picked('approved').filter((row) => missing.includes(row));
    const rejected = picked('rejected').filter((row) => missing.includes(row));
    const stored = db.prepare(`SELECT message_id FROM workflow_runs WHERE id = ?`).get(runId) as
      | { message_id: string }
      | undefined;
    // Same words as the Log's other rows; the Message-ID is only a fallback for a run whose label
    // predates this column.
    const mail = getMailLabel(db, runId) || stored?.message_id || '(unknown mail)';

    const written = await writeBack(budget, async () => {
      await book.markResults([
        ...approved.map((row) => ({ row, text: 'recorded ✓' })),
        ...rejected.map((row) => ({ row, text: 'cleared' })),
      ]);
      await book.appendLog([
        ...approved.map((row) => ({ when: stamp(), mail, did: `row ${row} · answered OK`, answer: 'OK', result: 'recorded ✓' })),
        ...rejected.map((row) => ({ when: stamp(), mail, did: `row ${row} · answered no`, answer: 'no', result: 'cleared' })),
      ]);
    });

    if (written) {
      // The sheet now shows the receipt, so the "not written yet" line it queued is stale.
      supersedeNotices(db, runId, NOT_WRITTEN_YET);
      audit(db, runId, 'pipeline', 'sheet-repaired');
      log(`repaired the sheet for ${mail} (${missing.length} row(s) the workbook had not seen)`);
    }
  }
}

/** Record what the human decided: in the row they answered on, in the Log sheet, and in the ledger. */
async function finalize(
  db: Db,
  book: Workbook,
  runId: string,
  review: Review,
  decisions: Decisions,
  budget?: { retries: number; pauseMs: number },
): Promise<Action> {
  const { approved, rejected, pending, unmatched } = decisions;
  const when = stamp();
  // The Log's `Mail` column is read by a shop owner, not by us: a Message-ID there is noise.
  const who = mailSource(review.mail) || review.messageId;

  // Not a purchase: nothing was drafted, so there is no decision to record — but the mail must
  // still leave a trace, or the person who sent it wonders where it went.
  if (review.intent === 'other') {
    setRunStatus(db, runId, 'ignored');
    audit(db, runId, 'pipeline', `intent=${review.intent}; not a purchase, nothing drafted`);
    const written = await writeBack(budget, () =>
      book.appendLog([
        { when, mail: who, did: 'not a purchase', answer: '—', result: 'ignored' },
      ]),
    );
    if (!written) audit(db, runId, 'pipeline', 'sheet-write-failed: workbook still open in Excel');
    return 'ignored';
  }

  for (const row of approved) audit(db, runId, 'human', `approved row ${row}`);
  for (const row of rejected) audit(db, runId, 'human', `rejected row ${row}`);

  if (review.rows.length === 0) {
    setRunStatus(db, runId, 'unmatched');
    audit(db, runId, 'pipeline', `no sheet row for [${unmatched.join(',')}]`);
    const written = await writeBack(budget, () =>
      book.appendLog([
        {
          when,
          mail: who,
          did: `no matching row: ${unmatched.join(', ') || 'nothing named'}`,
          answer: '—',
          result: 'not in sheet',
        },
      ]),
    );
    if (!written) audit(db, runId, 'pipeline', 'sheet-write-failed: workbook still open in Excel');
    return 'unmatched';
  }

  // The decision is recorded in the ledger first: holding it hostage to a file lock would lose it.
  const status = statusFor(decisions);
  setRunStatus(db, runId, status);
  if (pending.length > 0) audit(db, runId, 'pipeline', `waiting on rows [${pending.join(',')}]`);

  // Then write it where the human can see it. If Excel still has the file open, say so — the run
  // still finishes, because the ledger already knows.
  const written = await writeBack(budget, async () => {
    // A rejection removes the suggestion: the row goes back to being plain stock data.
    for (const row of rejected) await book.clearDraft(row);

    // The person who typed OK should see that it landed. Plain words, in their own row.
    await book.markResults([
      ...approved.map((row) => ({ row, text: 'recorded ✓' })),
      ...rejected.map((row) => ({ row, text: 'cleared' })),
    ]);

    // The working columns are transient — a rejection clears its draft. The Log sheet is the history.
    await book.appendLog(
      review.rows.map((entry) => ({
        when,
        mail: who,
        // The draft already carries the mail's own words, so repeating `matched` here would say
        // "row 4 · steel bracket · 30 · steel bracket".
        did: `row ${entry.row} · ${entry.draft}`,
        answer: approved.includes(entry.row) ? 'OK' : rejected.includes(entry.row) ? 'no' : '—',
        result: approved.includes(entry.row)
          ? 'recorded ✓'
          : rejected.includes(entry.row)
            ? 'cleared'
            : 'waiting',
      })),
    );
  });
  if (!written) {
    audit(db, runId, 'pipeline', 'sheet-write-failed: workbook still open in Excel');
    // The decision is safe in the ledger, but the person who typed OK must not be left wondering.
    notice(db, {
      runId,
      messageId: review.messageId,
      when,
      mail: who,
      did: `your answer on row(s) ${review.rows.map((entry) => entry.row).join(', ')} is recorded`,
      result: NOT_WRITTEN_YET,
    });
  }

  return status as Action;
}

/** Read every answer the human gave on these rows. A locked file is not an answer, so it throws. */
async function readDecisions(book: Workbook, rows: number[]): Promise<Decided[]> {
  const decided: Decided[] = [];
  for (const row of rows) {
    const answer = await book.readApproval(row);
    if (answer !== 'pending') decided.push({ row, decision: answer });
  }
  return decided;
}

async function resumeRun(
  workflow: Workflow,
  book: Workbook,
  db: Db,
  runId: string,
  review: Review,
  budget?: { retries: number; pauseMs: number },
): Promise<Action> {
  // The answers travel into the graph as data (the pattern every framework documents): the run
  // itself reads no file, so a workbook Excel is holding can no longer break a resume half-way.
  let decisions: Decided[];
  try {
    decisions = await readDecisions(
      book,
      review.rows.map((entry) => entry.row),
    );
  } catch (error) {
    if (!isSheetBusy(error)) throw error;
    audit(db, runId, 'pipeline', 'resume deferred: the sheet could not be read (open in Excel)');
    return 'deferred';
  }

  const run: Run = await workflow.createRun({ runId });
  const resumed = await run.resume({ step: 'review', resumeData: { decisions } });
  if (resumed.status !== 'success') {
    setRunStatus(db, runId, 'failed');
    audit(db, runId, 'pipeline', `resume-${resumed.status}`);
    return 'failed';
  }
  return finalize(db, book, runId, review, resumed.result, budget);
}

/**
 * One pass: finish runs that are waiting on a human, then take in new mail.
 * Safe to call on a timer — already-seen messages cost no model call.
 */
export async function pollInbox(options: InboxOptions): Promise<MailOutcome[]> {
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>((done) => setTimeout(done, ms)));
  const wait: Wait = {
    pollMs: options.pollMs ?? Number(process.env.SHEETFLOW_POLL_MS ?? 2_000),
    timeoutMs: options.timeoutMs ?? Number(process.env.SHEETFLOW_REVIEW_TIMEOUT_MS ?? 600_000),
    sleep,
  };
  const log = options.log ?? (() => {});

  // Make sure the human can see where to answer before we go looking for mail. A workbook that is
  // open in Excel cannot be refreshed — but reading it still works, so the pass goes ahead and only
  // the writes get held back until the file is free.
  try {
    await options.book.ensureInterface();
  } catch (error) {
    if (!isSheetBusy(error)) throw error;
    log('the workbook is open in Excel — close it and run again so the answering columns appear');
    // The owner will not see that log line. Leave it in the sheet for the next pass that gets in.
    notice(options.db, {
      runId: '(pass)',
      messageId: '(the workbook)',
      when: stamp(),
      mail: '(sheetflow itself)',
      did: 'could not create the answering columns — the workbook was open in Excel',
      result: 'close Excel and it will appear',
    });
  }

  // The workbook may still owe receipts, and the owner may still owe an explanation, from a pass
  // that ended while Excel held the file.
  await repairSheet(options.db, options.book, options.writeBack, log);
  await projectNotices(options.db, options.book, options.writeBack, log);

  // One deadline for the whole pass, not one per mail: with a ten-minute timeout and eleven mails a
  // first run would sit there for two hours. Whichever rows the human has marked by then are
  // resumed; the rest stay suspended on disk for the next run.
  const deadline = Date.now() + wait.timeoutMs;
  const remaining = (): Wait => ({
    ...wait,
    timeoutMs: Math.max(0, deadline - Date.now()),
  });

  const waiting = await resumeWaiting(options, { ...wait, timeoutMs: 0 }, log);
  const fresh: MailOutcome[] = [];
  for (const mail of await options.mailbox.fetch()) {
    fresh.push(await runOne(mail, options, remaining(), log));
  }
  // Anything this pass failed at, while the file was still writable.
  await projectNotices(options.db, options.book, options.writeBack, log);
  return [...waiting, ...fresh];
}

/**
 * Give up on runs that can never finish: still waiting when the workbook was replaced or the human
 * moved on, or failed on something temporary (Excel held the file). Forgetting the run is what makes
 * the same mail eligible again — the dedupe lives on that row.
 */
export async function forget(db: Db, statuses: RunStatus[], reason: string): Promise<string[]> {
  const forgotten: string[] = [];
  for (const status of statuses) {
    for (const row of listRunsByStatus(db, status)) {
      audit(db, row.id, 'pipeline', `abandoned (${status}): ${reason}`);
      deleteRun(db, row.id);
      forgotten.push(row.message_id);
    }
  }
  return forgotten;
}

/** Kept for the older name: runs left waiting on a human. */
export async function forgetWaiting(db: Db): Promise<string[]> {
  return forget(db, ['pending'], 'nobody answered');
}

/**
 * Runs left suspended by an earlier pass. The snapshot is on disk, so this works after a restart —
 * that is the whole point of `src/state.ts`.
 */
async function resumeWaiting(
  options: InboxOptions,
  wait: Wait,
  log: (line: string) => void,
): Promise<MailOutcome[]> {
  const { db } = options;
  const outcomes: MailOutcome[] = [];

  for (const row of listRunsByStatus(db, 'pending')) {
    const state = await options.workflow.getWorkflowRunById(row.id);
    // Gone: nothing to replay, and re-running the graph would draft twice.
    if (!state) continue;
    // Not suspended any more: the snapshot cannot be resumed, so the ledger must say so rather than
    // leave a row that looks like it is still waiting on a human who has already answered.
    if (state.status !== 'suspended') {
      setRunStatus(db, row.id, 'failed');
      audit(db, row.id, 'pipeline', `saved run is ${state.status}; nothing left to resume`);
      log(`failed ${row.message_id} — its saved run is ${state.status}; \`npm run inbox -- retry-failed\` drafts it again`);
      notice(db, {
        runId: row.id,
        messageId: row.message_id,
        when: stamp(),
        mail: getMailLabel(db, row.id) || row.message_id,
        did: `nothing to resume (its saved run is ${state.status})`,
        result: 'not drafted',
      });
      outcomes.push({ messageId: row.message_id, action: 'failed', rows: [], unmatched: [] });
      continue;
    }

    const review = suspendedReview(stepsOf(state)?.review?.suspendPayload);
    if (!review) {
      setRunStatus(db, row.id, 'failed');
      audit(db, row.id, 'pipeline', 'suspended without a review payload');
      notice(db, {
        runId: row.id,
        messageId: row.message_id,
        when: stamp(),
        mail: getMailLabel(db, row.id) || row.message_id,
        did: 'the saved run is unreadable',
        result: 'not drafted',
      });
      outcomes.push({ messageId: row.message_id, action: 'failed', rows: [], unmatched: [] });
      continue;
    }

    const rows = review.rows.map((entry) => entry.row);
    // A run left waiting by an earlier pass must never hold up new mail. Read the sheet once: if the
    // human has answered, finish the run; if not, say so and move on. (Measured 2026-10-28: five
    // undecided runs from a previous session silently ate the whole 10-minute budget before the pass
    // even fetched the mailbox — the command looked hung and wrote nothing.)
    if (review.intent === 'order' && rows.length > 0 && !(await waitForDecisions(options.book, rows, wait))) {
      audit(db, row.id, 'pipeline', `still undecided; left suspended (row(s) ${rows.join(',')})`);
      log(
        `waiting ${row.message_id} — row(s) ${rows.join(', ')} have no answer yet; ` +
          'open the workbook, pick OK or no in "Your answer", save and close',
      );
      outcomes.push({ messageId: row.message_id, action: 'pending', rows, unmatched: review.unmatched });
      continue;
    }

    const action = await resumeRun(options.workflow, options.book, db, row.id, review, options.writeBack);
    if (action === 'deferred') {
      log(`deferred ${row.message_id} — the workbook is open in Excel; save and close it and run again`);
    } else {
      log(`${action} ${row.message_id} (resumed)`);
    }
    outcomes.push({ messageId: row.message_id, action, rows, unmatched: review.unmatched });
  }

  return outcomes;
}

async function runOne(
  mail: Mail,
  options: InboxOptions,
  wait: Wait,
  log: (line: string) => void,
): Promise<MailOutcome> {
  const { db } = options;

  // The run id is minted first so the audit trail and the Mastra run share one identifier. A
  // duplicate message leaves this run unstarted — in memory only, nothing persisted.
  const run: Run = await options.workflow.createRun();
  const at = new Date().toISOString();
  const created = upsertRun(db, {
    id: run.runId,
    // A mail is its own conversation; there is no thread to join yet.
    thread_id: mail.messageId,
    message_id: mail.messageId,
    status: 'pending',
    created_at: at,
    updated_at: at,
  });

  if (!created) {
    log(`skip ${mail.messageId} (already handled)`);
    return { messageId: mail.messageId, action: 'skipped', rows: [], unmatched: [] };
  }

  // The workbook and the Log both speak in this label; the Message-ID stays in the ledger.
  setMailLabel(db, run.runId, mailSource(mail));

  const started = await run.start({
    inputData: {
      messageId: mail.messageId,
      from: mail.from,
      subject: mail.subject,
      body: mail.body,
      receivedAt: mail.receivedAt,
    },
  }).catch((error: unknown) => {
    if (isSheetBusy(error)) {
      // Nothing was written and nothing was decided: forget the ledger row so the next pass retries
      // this message instead of deduping it away.
      deleteRun(db, run.runId);
      log(`deferred ${mail.messageId}: the workbook is open in Excel — save and close it, then run again`);
      notice(db, {
        runId: run.runId,
        messageId: mail.messageId,
        when: stamp(),
        mail: mailSource(mail),
        did: BLOCKED_BY_EXCEL,
        result: 'retried automatically',
      });
      return undefined;
    }
    setRunStatus(db, run.runId, 'failed');
    const reason = reasonText(error);
    audit(db, run.runId, 'pipeline', `start-error: ${reason || 'unknown'}`);
    log(`fail ${mail.messageId} (${reason || 'unknown'})`);
    notice(db, { runId: run.runId, when: stamp(), ...failureNotice(reason, mail) });
    return undefined;
  });

  if (!started) {
    const deferred = getRun(db, mail.messageId) === undefined;
    return { messageId: mail.messageId, action: deferred ? 'deferred' : 'failed', rows: [], unmatched: [] };
  }

  const review = started.status === 'suspended' ? suspendedReview(stepsOf(started)?.review?.suspendPayload) : undefined;
  if (started.status === 'failed' && isSheetBusy(started.error)) {
    // Mastra reports a failed step as a result rather than throwing. Nothing was written and nothing
    // was decided, so forget the ledger row: the next pass must retry this message.
    deleteRun(db, run.runId);
    log(`deferred ${mail.messageId}: the workbook is open in Excel — save and close it, then run again`);
    notice(db, {
      runId: run.runId,
      messageId: mail.messageId,
      when: stamp(),
      mail: mailSource(mail),
      did: BLOCKED_BY_EXCEL,
      result: 'retried automatically',
    });
    return { messageId: mail.messageId, action: 'deferred', rows: [], unmatched: [] };
  }
  if (started.status !== 'suspended' || !review) {
    setRunStatus(db, run.runId, 'failed');
    const detail = reasonText((started as { error?: unknown }).error);
    // The technical text belongs in the audit trail; the owner gets a sentence he can act on.
    audit(db, run.runId, 'pipeline', `start-${started.status}`);
    if (detail) audit(db, run.runId, 'pipeline', `start-error: ${detail}`);
    log(`fail ${mail.messageId} (${started.status})`);
    notice(db, { runId: run.runId, when: stamp(), ...failureNotice(detail, mail) });
    return { messageId: mail.messageId, action: 'failed', rows: [], unmatched: [] };
  }

  const rows = review.rows.map((entry) => entry.row);
  // The model's own label is recorded only when it disagreed with the rule — that disagreement is
  // the early warning that the prompt and the fixtures have drifted apart.
  const modelLabel =
    review.modelIntent && review.modelIntent !== review.intent ? ` model=${review.modelIntent}` : '';
  audit(
    db,
    run.runId,
    'pipeline',
    `intent=${review.intent}${modelLabel} drafted rows [${rows.join(',')}] ` +
      `unmatched [${review.unmatched.join(',')}] no-quantity [${review.unquantified.join(',')}]`,
  );

  // Nothing to mark: not a purchase, no SKU matched, or no quantity. The run can finish right away.
  if (review.intent === 'other' || rows.length === 0) {
    // Built in 2026-10-28: a mail with an injected "SYSTEM NOTICE" made Nova Lite return
    // {"intent":"other"} with no lines at all. Nothing was obeyed and nothing leaked, but a real
    // order would have vanished as a plain question. Say so in the record.
    if (rows.length === 0 && QUANTITY_HINT.test(mail.body)) {
      audit(db, run.runId, 'pipeline', 'quantities in the text but no lines extracted — inspect this mail');
    }
    const action = await resumeRun(options.workflow, options.book, db, run.runId, review, options.writeBack);
    log(`${action} ${mail.messageId} (nothing to review)`);
    return { messageId: mail.messageId, action, rows, unmatched: review.unmatched };
  }

  // Say this before blocking: a command that sits silent for ten minutes reads as "hung".
  log(
    `drafted ${mail.messageId} into row(s) ${rows.join(', ')} — open the workbook, use the ` +
      `"Your answer: OK or no" column (pick OK or no), save and close it; ` +
      `waiting up to ${Math.round(wait.timeoutMs / 1000)}s`,
  );

  if (!(await waitForDecisions(options.book, rows, wait))) {
    // Left suspended on purpose: the human has not finished. The next pass picks it up.
    audit(db, run.runId, 'pipeline', `still undecided after ${wait.timeoutMs}ms; left suspended`);
    log(`waiting ${mail.messageId} (rows ${rows.join(',')} not all marked)`);
    return { messageId: mail.messageId, action: 'pending', rows, unmatched: review.unmatched };
  }

  const action = await resumeRun(options.workflow, options.book, db, run.runId, review, options.writeBack);
  log(`${action} ${mail.messageId} (status ${getRun(db, mail.messageId)?.status})`);
  return { messageId: mail.messageId, action, rows, unmatched: review.unmatched };
}



