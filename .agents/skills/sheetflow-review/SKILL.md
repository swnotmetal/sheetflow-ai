---
name: sheetflow-review
description: Rules for editing the sheetflow-ai pipeline. Use when changing workflow/model/extract/excel/mail/outbound/db code.
---
# sheetflow-ai review rules

Before editing `src/workflow.ts`, `src/extract.ts`, `src/model.ts`, `src/excel.ts`, `src/mail.ts`,
`src/outbound.ts` or `src/db.ts`:

1. A human decides before anything is applied. No model output reaches a customer, and no workbook
   cell is committed, without an explicit human mark (`AI_Review` = `OK`).
2. The model is explicit (`LLM_PROFILE`); nothing selects one implicitly. Fall back **only** on
   transport/availability errors, never on a validation or schema failure.
3. Values not literally present in the source are dropped (`groundFacts`). Never widen grounding to
   make a demo pass.
4. Business data stays in the workbook. The AI writes only to `AI_Draft` / `AI_Review`; existing
   cells are never rewritten. SQLite holds pipeline state and the audit trail only.
5. `test/*.test.ts` make no live model calls. Live checks live in `test/live.aws.ts`
   (`npm run test:live`) and are run deliberately.
6. Do not edit `d:\codes\Langgraph project` from this repository. It is a reference, not a source.
7. **Documentation budget: exactly two documents.** `README.md` (short, professional, readable by
   someone who does not know the code) and `HANDOVER.md` (local, gitignored). Never add a third:
   no ARCHITECTURE.md, no docs/ folder, no CHANGELOG. An explanation belongs in the README or a
   code comment — over-documentation is a defect here, not diligence.
8. Do not reinvent the wheel, always actively enable web search in github and find similiar projects and look into how they tackled the problem.
