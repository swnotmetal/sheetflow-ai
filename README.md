# sheetflow-ai

AI reads and writes **your existing Excel file** — without you leaving it. A lightweight,
SME-friendly sibling of [`langorder-ai`](https://github.com/swnotmetal/langorder-ai) (the
enterprise LangGraph + ERPNext variant). Where the enterprise version proves durable, recoverable
orchestration, this one proves the same safety story with **zero new infrastructure**: the
spreadsheet is the database and the UI.

## The idea

> Your orders live in a workbook. sheetflow drafts the extracted line items back into a column you
> already use, then waits. You approve by marking the row the way you always do (a colour, a tick,
> "OK"). Nothing is ever sent until a human says so.

## Stack

| Layer | Choice | Why |
| --- | --- | --- |
| Orchestration | [Mastra](https://mastra.ai) (TypeScript) | `suspend`/`resume` models "wait for the human" as a first-class primitive, not a patch |
| Model | multi-provider via `LLM_PROFILE` | `aws` · `openai` · `anthropic` · `groq` · `deepseek` · `local` (Ollama) — one prompt, one schema, one acceptance rule |
| Excel I/O | duduExcel MCP | atomic writes, format semantics preserved, `revert_last_write` |
| State + audit | `better-sqlite3` | `workflow_runs` + append-only `audit_log`; business data stays in the workbook |
| UI | CLI / TUI (Ink) | the review queue is a terminal command; Excel itself is the primary review surface |

## The one rule that never changes

A model's output is **never sent before a human approves it**, regardless of which model produced
it. Values the model invents but that are not in the source are dropped (`src/model.ts` grounds every
fact). This is the invariant carried over from `langorder-ai`.

## Run

```powershell
Copy-Item .env.example .env   # then set LLM_PROFILE
npm install
npm run dev
```

## Status

Skeleton. The workflow (`src/workflow.ts`) currently stubs the extract → draft → approve → send
pipeline; wiring the live Mastra `suspend/resume` and the duduExcel MCP client is the next step.
