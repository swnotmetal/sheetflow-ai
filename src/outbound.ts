// Where an approved draft goes.
//
// The whole project rests on one rule: nothing leaves the machine until a human approves it. That
// rule is enforced in `src/workflow.ts`, but this seam is what makes "sent" mean something — the
// default refuses to pretend, so a demo can never show a delivery that did not happen.

export interface OutboundLine {
  description: string;
  quantity?: string | null;
  unit?: string | null;
}

export interface OutboundMessage {
  messageId: string;
  row: number;
  lines: OutboundLine[];
  /** The model that produced the extraction behind this draft. */
  model: string;
}

export interface SendReceipt {
  to: string;
  at: string;
}

export interface Outbound {
  /** Deliver an approved draft. Callers must have a human `approve: true` in hand. */
  send(message: OutboundMessage): Promise<SendReceipt>;
}

export const outbound: Outbound = {
  async send() {
    throw new Error(
      'outbound.send: no outbound transport wired yet — refusing to report a delivery that did not happen',
    );
  },
};
