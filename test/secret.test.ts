// The mailbox password must not sit in the clear in a file the owner can open.
//
// These tests touch the real Windows DPAPI — no mock, because the whole value of the feature is that
// Windows, not this code, holds the key. On a non-Windows machine the round trip is skipped and only
// the honest refusal is asserted.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { canProtect, promptForSecret, protectSecret, revealSecret } from '../src/secret.js';
import { CONFIG_NAME, applyConfig, readConfig, storePassword } from '../src/config.js';

const windows = process.platform === 'win32';

test('a password survives the round trip without ever being stored in the clear', { skip: !windows }, async () => {
  const blob = await protectSecret('app-specific-password-42');
  assert.match(blob, /^[A-Za-z0-9+/=]+$/, 'a base64 blob, not the password');
  assert.doesNotMatch(blob, /app-specific-password-42/);
  assert.equal(await revealSecret(blob), 'app-specific-password-42');
});

test('a blob cannot be decrypted once it is not a blob', { skip: !windows }, async () => {
  await assert.rejects(revealSecret('bm90LWEtYmxvYg=='), /could not read the stored mailbox password/);
});

test('set-password refuses a pipe instead of hanging on it', { skip: !windows || process.stdin.isTTY }, async () => {
  // Found the hard way: with stdin redirected, PowerShell's secure prompt waits for a console that
  // will never arrive and the command sits there forever. Refusing up front is the only kind answer.
  await assert.rejects(promptForSecret(), /needs a console/);
});

test('elsewhere the tool refuses instead of connecting with no password', { skip: windows }, async () => {
  await assert.rejects(protectSecret('x'), /only supported on Windows/);
  await assert.rejects(revealSecret('x'), /cannot be read on this system/);
});

test('set-password replaces the plaintext field and keeps the rest of the file', { skip: !windows }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-pw-'));
  try {
    const original = {
      workbook: 'orders.xlsx',
      llmProfile: 'openai',
      imap: { host: 'imap.example.com', user: 'shop@example.com', password: 'typed-by-hand' },
    };
    await writeFile(join(dir, CONFIG_NAME), JSON.stringify(original), 'utf8');

    const blob = await protectSecret('real-one');
    await storePassword(blob, dir);

    const onDisk = await readFile(join(dir, CONFIG_NAME), 'utf8');
    assert.doesNotMatch(onDisk, /typed-by-hand/, 'the plaintext password is gone');
    const saved = readConfig(dir)!;
    assert.equal(saved.workbook, 'orders.xlsx', 'the owner’s other answers are untouched');
    assert.equal(saved.imap?.host, 'imap.example.com');
    assert.equal(saved.imap?.user, 'shop@example.com');
    assert.equal(saved.imap?.password, undefined);
    assert.equal(saved.imap?.passwordProtected, blob);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('the run decrypts the stored password into the environment, once', { skip: !windows }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-pw2-'));
  const saved = { ...process.env };
  try {
    await writeFile(
      join(dir, CONFIG_NAME),
      JSON.stringify({ imap: { host: 'imap.example.com', passwordProtected: await protectSecret('s3cret') } }),
      'utf8',
    );
    delete process.env.IMAP_PASS;

    await applyConfig(dir);

    assert.equal(process.env.IMAP_PASS, 's3cret');
    assert.equal(process.env.IMAP_HOST, 'imap.example.com');
  } finally {
    process.env = saved;
    await rm(dir, { recursive: true, force: true });
  }
});

test('an environment password outranks the stored one', { skip: !windows }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sheetflow-pw3-'));
  const saved = { ...process.env };
  try {
    await writeFile(
      join(dir, CONFIG_NAME),
      JSON.stringify({ imap: { host: 'imap.example.com', passwordProtected: await protectSecret('from-file') } }),
      'utf8',
    );
    process.env.IMAP_PASS = 'from-shell';

    await applyConfig(dir);

    assert.equal(process.env.IMAP_PASS, 'from-shell', 'a developer’s shell is not overridden by a file');
  } finally {
    process.env = saved;
    await rm(dir, { recursive: true, force: true });
  }
});

test('the capability check matches the platform', () => {
  assert.equal(canProtect(), windows);
});
