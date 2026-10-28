// The pipeline: mail -> extract -> match SKUs to sheet rows -> draft into AI_Draft -> suspend for
// a human -> read the human's mark in the sheet.
//
// This is why the project chose Mastra: `suspend()`/`resume()` model "wait for the human" as a
// first-class primitive. The run literally stops, and Mastra keeps the suspended snapshot until
// the sheet says what the human decided.
//
// The model and the workbook are seams (`src/extract.ts`, `src/excel.ts`); `buildWorkflow` takes
// them as arguments so a test can run the real graph without a network or a real mailbox.
//
// There is deliberately no outbound step: an approved draft is recorded, not emailed. `src/outbound.ts`
// keeps the refusal-to-pretend seam for a sandbox reply later.

import { createStep, createWorkflow } from '@mastra/core/workflows';
import type { Mastra } from '@mastra/core';
import { z } from 'zod';
import { resolveProfile, INTENTS, type Intent, type ResolvedProvider } from './model.js';
import { extractFacts, type ExtractionCall } from './extract.js';
import { workbook, type Workbook } from './excel.js';
import { durableMastra } from './state.js';

// The mail itself is the input. `subject` and `from` are part of the grounding text: a customer PO
// number very often exists only in the subject, and a sender must never be invented.
const TriggerSchema = z.object({
  messageId: z.string(),
  from: z.string().default(''),
  subject: z.string().default(''),
  body: z.string(),
});

/** One human answer, as data. The sheet collects it; the driver hands it back through `resume`.
 * The shape follows the frameworks' own examples (LangGraph `Command(resume=...)`, Mastra
 * `resumeData`): the decision is an input to the graph, not something the graph goes and reads.
 * That is what keeps a resumed run replayable, and it is why a locked workbook can no longer break
 * a resume half-way through.
 */
const DecisionSchema = z.object({
  row: z.number().int(),
  decision: z.enum(['approved', 'rejected']),
});

/** What the model may cite back as literally present in the mail. */
export function groundText(mail: { from: string; subject: string; body: string }): string {
  return [mail.subject, mail.body, mail.from].filter((part) => part.trim()).join('\n');
}

const LineSchema = z.object({
  description: z.string(),
  sku: z.string().nullable().optional(),
  quantity: z.string().nullable().optional(),
  unit: z.string().nullable().optional(),
});

const ExtractOutput = z.object({
  messageId: z.string(),
  /** What the mail is. Only an `order` is allowed to reach the sheet. */
  intent: z.enum(INTENTS),
  /** Grounded lines only — anything the model invented was already dropped. */
  lines: z.array(LineSchema),
  dropped: z.number().int(),
  // Which model actually answered. Not always `LLM_PROFILE`'s first choice: a transport failure
  // can move to the profile's fallback, and the audit trail has to say so.
  model: z.string(),
});

/** One drafted sheet row. */
const DraftRowSchema = z.object({
  row: z.number().int(),
  /** The SKU or product name the mail named, as written. */
  matched: z.string(),
  /** The value the AI proposes for that row's `AI_Draft` cell. */
  draft: z.string(),
});

const DraftOutput = z.object({
  messageId: z.string(),
  intent: z.enum(INTENTS),
  /** What the model called it. Kept so a disagreement with the rule below stays visible. */
  modelIntent: z.enum(INTENTS),
  model: z.string(),
  rows: z.array(DraftRowSchema),
  /** Lines the sheet has no row for. Reported, never written. */
  unmatched: z.array(z.string()),
  /** Lines with no quantity — nothing to move, so nothing is drafted. */
  unquantified: z.array(z.string()),
  dropped: z.number().int(),
});

const ReviewOutput = z.object({
  intent: z.enum(INTENTS),
  model: z.string(),
  approved: z.array(z.number().int()),
  rejected: z.array(z.number().int()),
  pending: z.array(z.number().int()),
  unmatched: z.array(z.string()),
});

/**
 * The model proposes, the text disposes — but only on what the *text* can settle.
 *
 * A quantity is text we can point at, so a mail containing one is an order even if the model called
 * it a question (Nova Lite did that 3 times in 10 fixtures). What the text cannot settle is why a
 * customer asked something: "do you have any black crew socks?" is not a question for this project
 * to interpret, it is a row a human should see. So the only thing `other` is allowed to suppress is
 * mail that is not about buying from us at all — recruitment, vendor pitches, newsletters.
 */
export function decideIntent(modelIntent: Intent, lines: { quantity?: string | null }[]): Intent {
  if (lines.some((line) => line.quantity?.trim())) return 'order';
  return modelIntent === 'other' ? 'other' : 'question';
}

/** What a human reads in the sheet before deciding: quantity + unit + what the mail named. */
export function formatDraft(lines: z.infer<typeof LineSchema>[]): string {
  return lines
    .map((line) => [line.quantity, line.unit, line.description].filter(Boolean).join(' '))
    .join('; ');
}

export interface WorkflowDeps {
  /** Injected in tests; production uses the real `.xlsx` client in `src/excel.ts`. */
  book?: Workbook;
  /** Injected in tests; production calls the profile's model. */
  extract?: ExtractionCall;
  /** Resolved per run from `LLM_PROFILE` unless a caller names it explicitly. */
  provider?: ResolvedProvider;
  /**
   * Mastra owns the snapshot store `resume()` reads from. Defaults to a file-backed store, so a run
   * suspended by one `npm run inbox` can be resumed by the next one — hours later, after a restart.
   * Tests pass their own instance.
   */
  mastra?: Mastra;
}

let defaultMastra: Mastra | undefined;

export function buildWorkflow(deps: WorkflowDeps = {}) {
  const book = deps.book ?? workbook;
  const mastra = deps.mastra ?? (defaultMastra ??= durableMastra());

  const extractStep = createStep({
    id: 'extract',
    inputSchema: TriggerSchema,
    outputSchema: ExtractOutput,
    execute: async ({ inputData }) => {
      const provider = deps.provider ?? resolveProfile(process.env.LLM_PROFILE);
      // The catalogue is the sender's own sheet: it is what makes "black crew socks" resolvable
      // to `ACC-SOC-BLK-OS`, and what stops the model naming a product we do not sell.
      const catalogue = await book.products();
      const { kept, dropped, intent, model } = await extractFacts(
        inputData.body,
        provider,
        deps.extract,
        groundText(inputData),
        catalogue,
      );

      return {
        messageId: inputData.messageId,
        intent,
        lines: kept.lines ?? [],
        dropped: dropped.length,
        model,
      };
    },
  });

  const draftStep = createStep({
    id: 'draft',
    inputSchema: ExtractOutput,
    outputSchema: DraftOutput,
    execute: async ({ inputData }) => {
      const intent = decideIntent(inputData.intent, inputData.lines);

      // Not a purchase at all: nothing is drafted, and nothing waits for a human. A recruitment
      // pitch has no business writing into a stock sheet. A *question* is different — the sender
      // named something we sell, so a human sees the row even without a quantity.
      if (intent === 'other') {
        return {
          messageId: inputData.messageId,
          intent,
          modelIntent: inputData.intent,
          model: inputData.model,
          rows: [],
          unmatched: inputData.lines.map((line) => line.description),
          unquantified: [],
          dropped: inputData.dropped,
        };
      }

      // The row a draft belongs to is decided by the sheet, not by the caller.
      const { rows, missing } = await book.resolveRows(inputData.lines);

      const unquantified = inputData.lines
        .filter((line) => !line.quantity?.trim())
        .map((line) => line.description);

      const drafts = new Map<number, { matched: string; draft: string }>();
      inputData.lines.forEach((line, at) => {
        const row = rows.get(at);
        if (!row) return;
        // No quantity in the mail means no quantity is written. `?` is the honest placeholder:
        // the human types the number (or `no`) in the answer column.
        const value = [line.quantity?.trim(), line.unit?.trim()].filter(Boolean).join(' ') || '?';
        const existing = drafts.get(row);
        // Two lines for one SKU stay visible as two proposals; summing quantities would invent a
        // number the mail never stated.
        drafts.set(row, {
          matched: line.description,
          draft: existing ? `${existing.draft} + ${value}` : value,
        });
      });

      const written: z.infer<typeof DraftRowSchema>[] = [];
      for (const [row, { matched, draft }] of drafts) {
        await book.writeDraft(draft, { row });
        written.push({ row, matched, draft });
      }

      return {
        messageId: inputData.messageId,
        intent,
        modelIntent: inputData.intent,
        model: inputData.model,
        rows: written,
        unmatched: missing,
        unquantified,
        dropped: inputData.dropped,
      };
    },
  });

  const reviewStep = createStep({
    id: 'review',
    inputSchema: DraftOutput,
    outputSchema: ReviewOutput,
    // What we hand to the human, and what they hand back: one answer per row they touched.
    suspendSchema: z.object({
      messageId: z.string(),
      intent: z.enum(INTENTS),
      modelIntent: z.enum(INTENTS),
      rows: z.array(DraftRowSchema),
      unmatched: z.array(z.string()),
      unquantified: z.array(z.string()),
    }),
    resumeSchema: z.object({ decisions: z.array(DecisionSchema) }),
    execute: async ({ inputData, suspend, resumeData }) => {
      if (!resumeData) {
        // First pass: the drafts are in the sheet; stop and wait. Nothing else happens yet.
        return suspend(
          {
            messageId: inputData.messageId,
            intent: inputData.intent,
            modelIntent: inputData.modelIntent,
            rows: inputData.rows,
            unmatched: inputData.unmatched,
            unquantified: inputData.unquantified,
          },
          { resumeLabel: 'sheet-review' },
        );
      }

      // Resumed: the answers arrive as data. A row the human left alone stays `pending` — silence is
      // not consent — and clearing a rejected draft is the driver's write-back, not this step's job.
      const answered = new Map(resumeData.decisions.map(({ row, decision }) => [row, decision]));
      const approved: number[] = [];
      const rejected: number[] = [];
      const pending: number[] = [];
      for (const { row } of inputData.rows) {
        const decision = answered.get(row);
        if (decision === 'approved') approved.push(row);
        else if (decision === 'rejected') rejected.push(row);
        else pending.push(row);
      }

      return { intent: inputData.intent, model: inputData.model, approved, rejected, pending, unmatched: inputData.unmatched };
    },
  });

  return createWorkflow({
    id: 'sheetflow-order',
    mastra,
    inputSchema: TriggerSchema,
    outputSchema: ReviewOutput,
  })
    .then(extractStep)
    .then(draftStep)
    .then(reviewStep)
    .commit();
}
