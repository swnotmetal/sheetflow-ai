// The file the shop owner edits instead of a wall of environment variables.
//
// What matters: a path in the file means "next to the file", not "next to wherever the shortcut
// started from"; a broken file refuses to start instead of silently pointing at the wrong workbook;
// and an explicitly set environment variable is never overridden by the file.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { CONFIG_NAME, envFromConfig, readConfig } from '../src/config.js';

async function configDir(contents: string) {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-config-'));
  await writeFile(join(dir, CONFIG_NAME), contents, 'utf8');
  return dir;
}

test('a missing config file is not an error — the defaults apply', () => {
  const dir = join(tmpdir(), 'sheetflow-definitely-not-here');
  assert.equal(readConfig(dir), undefined);
});

test('the file says which workbook, mailbox and model', async () => {
  const dir = await configDir(
    JSON.stringify({
      workbook: 'my stock.xlsx',
      llmProfile: 'openai',
      llmApiKey: 'sk-not-a-real-key',
      imap: { host: 'imap.example.com', port: 993, user: 'shop@example.com', password: 'x' },
      reviewTimeoutMs: 0,
    }),
  );
  try {
    const env = envFromConfig(readConfig(dir)!, dir);
    assert.equal(env.WORKBOOK_PATH, join(dir, 'my stock.xlsx'), 'relative to the file, not the cwd');
    assert.equal(env.LLM_PROFILE, 'openai');
    assert.equal(env.IMAP_HOST, 'imap.example.com');
    assert.equal(env.IMAP_PORT, '993');
    assert.equal(env.SHEETFLOW_REVIEW_TIMEOUT_MS, '0', '0 is a value, not "unset"');
    assert.equal(env.SHEETFLOW_DATA_DIR, join(dir, '.data'), 'the ledger lives beside the config');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('an absolute path is left exactly as written', async () => {
  const dir = await configDir(JSON.stringify({ workbook: 'D:\\stock\\orders.xlsx' }));
  try {
    assert.equal(envFromConfig(readConfig(dir)!, dir).WORKBOOK_PATH, 'D:\\stock\\orders.xlsx');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('a file with nothing in it still points at a workbook beside itself', async () => {
  const dir = await configDir('{}');
  try {
    const env = envFromConfig(readConfig(dir)!, dir);
    assert.equal(env.WORKBOOK_PATH, join(dir, 'orders.xlsx'));
    assert.equal(env.MAIL_DIR, join(dir, 'mail'));
    assert.equal(env.IMAP_HOST, undefined, 'no mailbox configured is not an empty host');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('a broken config file refuses to start, in one sentence', async () => {
  const dir = await configDir('{ "workbook": "orders.xlsx", ');
  try {
    assert.throws(() => readConfig(dir), /is not valid JSON[\s\S]*Fix it, or delete it/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('a config file that is not an object is rejected', async () => {
  const dir = await configDir('["orders.xlsx"]');
  try {
    assert.throws(() => readConfig(dir), /single \{ … \} object/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
