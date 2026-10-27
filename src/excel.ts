// The only thing sheetflow touches is the user's workbook, and it touches it through an MCP server
// (duduExcel) rather than hand-rolled `xlsx` glue. That buys three guarantees we would otherwise
// have to build ourselves:
//   - atomic writes (temp file -> replace, with a .bak) and `revert_last_write`
//   - format semantics are preserved (strikethrough = cancelled, highlight = needs review)
//   - server-side aggregation, so a 5000-row sheet never floods the context window
//
// This module is the seam. The concrete MCP client is wired in the workflow; tests inject a
// `Workbook` implementation instead of talking to a real file.

export interface DraftCell {
  row: number;
  column: string; // e.g. "AI_Draft" or "J"
}

export interface Workbook {
  /** Append a draft value to a column, preserving everything else in the file. */
  writeDraft(value: string, cell: DraftCell): Promise<void>;
  /** Read the human's approval marker for a row (colour / tick / "OK"). */
  readApproval(row: number): Promise<'approved' | 'rejected' | 'pending'>;
  /** Undo the most recent write. The safety net for "the AI edited my sheet". */
  revertLastWrite(): Promise<void>;
}

// A clearly-labelled stub until the duduExcel MCP client is wired. Everything that matters is
// behind `Workbook`, so the rest of the pipeline does not care how the sheet is opened.
export const workbook: Workbook = {
  async writeDraft() {
    throw new Error('workbook.writeDraft: duduExcel MCP client not wired yet');
  },
  async readApproval() {
    throw new Error('workbook.readApproval: duduExcel MCP client not wired yet');
  },
  async revertLastWrite() {
    throw new Error('workbook.revertLastWrite: duduExcel MCP client not wired yet');
  },
};
