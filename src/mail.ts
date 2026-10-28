// Where the mail comes from.
//
// **Read-only.** No flag is set, nothing is moved, nothing is deleted: if the pipeline stops, the
// mailbox is exactly as the operator left it. That also means the same message is fetched again on
// the next poll — dedupe is the database's job (`workflow_runs.message_id` is UNIQUE), not the
// mailbox's.
//
// The MIME decode is deliberately the same shape as the sibling repo's `src/mime.ts`: headers are
// metadata, the body is the text a human wrote, and the subject is kept separate because a
// customer PO number very often exists only there.

import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

export interface Mail {
  /** Stable dedupe key: the `Message-ID` header, or a UID-derived fallback. */
  messageId: string;
  from: string;
  subject: string;
  /** Text a human wrote. Headers are excluded. */
  body: string;
  receivedAt: string;
}

export interface Mailbox {
  /** Look at the mailbox without changing it. */
  fetch(limit?: number): Promise<Mail[]>;
}

function textFromHtml(html: string | false | undefined): string {
  if (typeof html !== 'string') return '';
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Decode one RFC822 message. Exported so tests can exercise the decoder without a server. */
export async function parseMail(raw: Buffer | string, fallbackId = ''): Promise<Mail> {
  const parsed = await simpleParser(raw);
  const from = parsed.from?.value?.[0]?.address ?? '';
  return {
    messageId: (parsed.messageId ?? fallbackId).trim() || fallbackId,
    from,
    subject: (parsed.subject ?? '').trim(),
    body: ((parsed.text ?? '') || textFromHtml(parsed.html)).trim(),
    receivedAt: (parsed.date ?? new Date()).toISOString(),
  };
}

/** A folder of `.eml` files. The demo path: no credentials, no network. */
export function fileMailbox(dir: string): Mailbox {
  return {
    async fetch(limit = 10) {
      let names: string[];
      try {
        names = await readdir(dir);
      } catch {
        return [];
      }
      const files = names.filter((name) => name.toLowerCase().endsWith('.eml')).sort().slice(-limit);
      const mails: Mail[] = [];
      for (const name of files) {
        mails.push(await parseMail(await readFile(join(dir, name)), name));
      }
      return mails;
    },
  };
}

function required(env: NodeJS.ProcessEnv, key: string): string {
  const value = env[key]?.trim();
  if (!value) throw new Error(`${key} is required for the IMAP mailbox (set it in .env)`);
  return value;
}

/**
 * A read-only IMAP client. Connection is opened per poll and closed again, so a crashed run cannot
 * leave a half-locked mailbox behind.
 */
export function imapMailbox(env: NodeJS.ProcessEnv = process.env): Mailbox {
  return {
    async fetch(limit = 10) {
      const client = new ImapFlow({
        host: required(env, 'IMAP_HOST'),
        port: Number(env.IMAP_PORT?.trim() || 993),
        secure: (env.IMAP_SECURE?.trim() || 'true') !== 'false',
        auth: { user: required(env, 'IMAP_USER'), pass: required(env, 'IMAP_PASS') },
        logger: false,
      });

      await client.connect();
      const lock = await client.getMailboxLock(env.IMAP_MAILBOX?.trim() || 'INBOX', {
        readOnly: true,
      });
      try {
        // Envelopes first so "the newest N" is a decision, not whatever the server streams first.
        const uids: number[] = [];
        for await (const message of client.fetch({ seen: false }, { uid: true })) {
          if (typeof message.uid === 'number') uids.push(message.uid);
        }
        const wanted = uids.slice(-limit);
        if (wanted.length === 0) return [];

        const mails: Mail[] = [];
        for await (const message of client.fetch(
          { uid: wanted.join(',') },
          { source: true, uid: true },
        )) {
          if (!message.source) continue;
          mails.push(await parseMail(message.source, `imap:${message.uid}`));
        }
        return mails;
      } finally {
        lock.release();
        await client.logout();
      }
    },
  };
}

/** Pick the mailbox the environment asks for: IMAP when configured, `.eml` folder otherwise. */
export function configuredMailbox(env: NodeJS.ProcessEnv = process.env): Mailbox {
  if (env.IMAP_HOST?.trim()) return imapMailbox(env);
  return fileMailbox(env.MAIL_DIR?.trim() || join('test', 'mail'));
}
