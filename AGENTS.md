# sheetflow-ai — design invariants

A lightweight sibling of `langorder-ai`. Same safety story, no ERP: the user's own spreadsheet is the
database and the interface. The model and the workbook are seams; the rules below are what the code
exists to protect. A change that breaks one of them is a regression even if every test passes.

## Invariants (do not regress)

1. **Nothing is applied without a human.** sheetflow drafts a quantity into a working column and
   stops; the human types `OK` or `no`. It never writes the final result before that answer, and it
   never sends email at all.
2. **The review surface is the user's own sheet.** The three columns (`AI suggests`,
   `Your answer: OK or no`, `AI result`) are the whole interface. Do not add a second screen.
3. **Every value must be traceable to the mail.** `groundFacts` drops anything not literally in the
   source. Non-English or complex product text is the highest-risk axis for hallucination, so
   grounding matters more there, not less.
4. **Never rewrite the user's sheet beyond the working columns.** Formulas, formats and unrelated
   cells are preserved; writes are atomic (temp → rename with `.bak`) and reversible.
5. **It replaces manual copy-paste, not the salesperson.** It moves rows from the mailbox into the
   sheet; it never contacts the customer. The higher the ticket, the more a human sits in between.

## Scope

High-volume, low-to-mid-ticket, repeatable orders — cheap, bulky, "same thing every week". Not every
deal, and not for high-touch or high-ticket sales.

**One editor, one local file.** The local `.xlsx` direction assumes the workbook has a single owner
and does **not** solve concurrent editing: if a colleague works in the same file they keep their own
numbering and colouring and stay out of the AI's three columns. That is a stated limit, not a gap to
close — a local file cannot reconcile two writers, and building for it would cost the simplicity
that makes the direction worth doing. A shared Google Sheet has no such problem, which is why
`src/sheets.ts` exists rather than a lock protocol for a file.

## Model access

One provider per run, chosen by `LLM_PROFILE` (`aws` · `openai` · `anthropic` · `groq` · `deepseek` ·
`local`). Never select a model implicitly. Tests make no live model calls.

## Evidence

Label mock, real and unverified claims precisely. Fictional data only.
