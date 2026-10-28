// MIME decoding and the `.eml` mailbox. No IMAP server, no network.

import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { fileMailbox, parseMail } from '../src/mail.js';

const FIXTURE = 'test/mail/01-all-in-stock.eml';

test('a real message header is metadata, and the body stays body', async () => {
  const mail = await parseMail(await readFile(FIXTURE), 'fallback');

  assert.equal(mail.messageId, '<mix-01@hanoi-threads.example>');
  assert.equal(mail.from, 'nguyen.van.a@hanoi-threads.example');
  assert.equal(mail.subject, 'RFQ: urgent stock check for next season');
  assert.match(mail.body, /40 units/);
  // Headers must not leak into the text a human wrote — grounding would then "cite" them.
  assert.doesNotMatch(mail.body, /Message-ID|Content-Type/i);
});

test('a mail with no Message-ID falls back to the caller-supplied id', async () => {
  const mail = await parseMail('Subject: hi\r\n\r\nbody only', 'imap:42');
  assert.equal(mail.messageId, 'imap:42');
  assert.equal(mail.body, 'body only');
});

test('an HTML-only mail is flattened to text', async () => {
  const raw = [
    'From: a@b.example',
    'Subject: html',
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset="utf-8"',
    '',
    '<p>Ship <b>5</b> units</p><script>ignore()</script>',
  ].join('\r\n');

  const mail = await parseMail(raw);
  assert.match(mail.body, /Ship 5 units/);
  assert.doesNotMatch(mail.body, /ignore/);
});

test('the file mailbox returns .eml files in name order, newest last', async () => {
  const all = await fileMailbox('test/mail').fetch(20);
  assert.equal(all.length, 11);
  assert.equal(all[0].subject, 'RFQ: urgent stock check for next season');

  // `limit` keeps the newest N: a poll should not be flooded by the oldest mail in the folder.
  const newest = await fileMailbox('test/mail').fetch(2);
  assert.deepEqual(
    newest.map((mail) => mail.subject),
    ['RFQ: priority stock allocation - Q4 project', 'quick question'],
  );
});

test('a missing mail directory is an empty inbox, not a crash', async () => {
  assert.deepEqual(await fileMailbox('test/does-not-exist').fetch(), []);
});
