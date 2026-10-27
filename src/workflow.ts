// The pipeline: extract -> draft into the workbook -> suspend for a human -> send on approval.
//
// This is the whole reason the project chose Mastra over LangGraph: `suspend()`/`resume()` model
// "wait for the human" as a first-class primitive. The step below literally stops mid-flow and
// returns, and Mastra persists that suspended state until someone calls `resume`.
//
// The model and the workbook are seams (`src/model.ts`, `src/excel.ts`); this file does not care
// which provider extracts or how the sheet is opened.

import { createStep, createWorkflow } from '@mastra/core/workflows';
import { z } from 'zod';
import { groundFacts, type Extraction } from './model.js';
import { workbook } from './excel.js';

const TriggerSchema = z.object({
  messageId: z.string(),
  body: z.string(),
  row: z.number().int(),
});

const extractStep = createStep({
  id: 'extract',
  inputSchema: TriggerSchema,
  outputSchema: z.object({
    messageId: z.string(),
    row: z.number().int(),
    lines: z.number().int(),
    dropped: z.number().int(),
  }),
  execute: async ({ inputData }) => {
    // TODO: call the model for `inputData.body` via the resolved LLM_PROFILE, then ground the
    // result. Until the transport is wired this returns an empty, fully-grounded extraction.
    const raw: Extraction = { lines: [] };
    const { kept, dropped } = groundFacts(raw, inputData.body);
    return {
      messageId: inputData.messageId,
      row: inputData.row,
      lines: kept.lines?.length ?? 0,
      dropped: dropped.length,
    };
  },
});

const reviewStep = createStep({
  id: 'review',
  inputSchema: z.object({
    messageId: z.string(),
    row: z.number().int(),
    lines: z.number().int(),
    dropped: z.number().int(),
  }),
  outputSchema: z.object({ sent: z.boolean() }),
  // What we hand to the human, and what they hand back.
  suspendSchema: z.object({ messageId: z.string(), row: z.number().int() }),
  resumeSchema: z.object({ approve: z.boolean() }),
  execute: async ({ inputData, suspend, resumeData }) => {
    if (!resumeData) {
      // First pass: write the draft into the workbook, then stop and wait. Nothing is sent yet.
      await workbook.writeDraft('(draft)', { row: inputData.row, column: 'AI_Draft' });
      return suspend({ messageId: inputData.messageId, row: inputData.row }, { resumeLabel: 'approve' });
    }
    // Resumed with the human's decision. Only `approve: true` sends — the invariant holds.
    if (resumeData.approve) {
      // TODO: send via the outbound guard (loopback/sandbox only, mirroring langorder-ai).
    }
    return { sent: resumeData.approve };
  },
});

export const workflow = createWorkflow({
  id: 'sheetflow-order',
  inputSchema: TriggerSchema,
  outputSchema: z.object({ sent: z.boolean() }),
})
  .then(extractStep)
  .then(reviewStep)
  .commit();
