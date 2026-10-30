// The mailbox password, kept where Windows keeps secrets.
//
// A shop owner has to give this tool a mailbox password, and a plaintext `password` field in
// `config.json` is the wrong answer twice over: it sits in a file anyone can open, and it invites the
// owner to type a real password into a program they cannot inspect. So the password is stored as a
// DPAPI blob — encrypted by Windows, for **this user on this machine only**. Copying the config to
// another PC does not carry the secret; it simply fails to decrypt.
//
// DPAPI is reached through PowerShell rather than a native Node module: a native module cannot be
// embedded in the single executable, and this is the one path that needs no extra dependency. The
// plaintext never travels in an argument (it would be visible in the process list) — it goes in on
// stdin, and comes back on stdout.
//
// The limit is honest: a DPAPI blob is unreadable on macOS or Linux, so `revealSecret` says exactly
// that instead of quietly connecting with no password.

import { spawn } from 'node:child_process';

/** Windows protects the blob for the current user; there is no equivalent here. */
export const canProtect = (): boolean => process.platform === 'win32';

const ENCODED_HEADER = `$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Security
`;

/** Read the plaintext on stdin, print the DPAPI blob on stdout. */
const PROTECT_SCRIPT = `${ENCODED_HEADER}$plain = [Console]::In.ReadToEnd()
if ([string]::IsNullOrEmpty($plain)) { throw 'nothing to protect' }
$bytes = [Text.Encoding]::UTF8.GetBytes($plain)
$prot = [Security.Cryptography.ProtectedData]::Protect($bytes, $null, 'CurrentUser')
[Console]::Out.Write([Convert]::ToBase64String($prot))
`;

/** Read the DPAPI blob on stdin, print the plaintext on stdout. */
const REVEAL_SCRIPT = `${ENCODED_HEADER}$blob = [Console]::In.ReadToEnd().Trim()
if ([string]::IsNullOrEmpty($blob)) { throw 'nothing to reveal' }
$bytes = [Convert]::FromBase64String($blob)
$plain = [Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, 'CurrentUser')
[Console]::Out.Write([Text.Encoding]::UTF8.GetString($plain))
`;

/** Ask for the password without echoing it, then print the blob. Needs a real console. */
const PROMPT_SCRIPT = `${ENCODED_HEADER}$secure = Read-Host -Prompt 'Mailbox password (not shown, not stored in plain text)' -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
if ([string]::IsNullOrEmpty($plain)) { throw 'nothing entered' }
$bytes = [Text.Encoding]::UTF8.GetBytes($plain)
$prot = [Security.Cryptography.ProtectedData]::Protect($bytes, $null, 'CurrentUser')
[Console]::Out.Write([Convert]::ToBase64String($prot))
`;

interface RunResult {
  stdout: string;
  stderr: string;
  code: number;
}

/**
 * Run one of the scripts above. `-EncodedCommand` takes UTF-16LE base64, which is why a script full
 * of quotes, braces and `$` needs no escaping at all.
 */
function runPowerShell(script: string, input: string, interactive = false): Promise<RunResult> {
  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  const args = ['-NoProfile', ...(interactive ? [] : ['-NonInteractive']), '-EncodedCommand', encoded];
  return new Promise((resolve, reject) => {
    const child = spawn('powershell.exe', args, {
      // Interactive mode must keep the console: `Read-Host` has nowhere to read from otherwise.
      stdio: [interactive ? 'inherit' : 'pipe', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let stdout = '';
    let stderr = '';
    child.stdout?.on('data', (chunk: Buffer) => (stdout += chunk.toString('utf8')));
    child.stderr?.on('data', (chunk: Buffer) => (stderr += chunk.toString('utf8')));
    child.on('error', reject);
    child.on('close', (code) => resolve({ stdout, stderr, code: code ?? 1 }));
    if (!interactive) {
      child.stdin?.end(input, 'utf8');
    }
  });
}

/** PowerShell wraps errors in CLIXML when its output is redirected; recover the sentence inside. */
function cleanError(result: RunResult): string {
  const text = (result.stderr || result.stdout).trim();
  const message = /<S S="Error">([\s\S]*?)<\/S>/.exec(text)?.[1] ?? text;
  const decoded = message
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    // CLIXML escapes characters that a message may not carry literally.
    .replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
    .trim();
  // The first line carries the exception type; the useful sentence is usually next to it.
  return (
    decoded
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line && !line.startsWith('+') && !line.startsWith('At line'))
      .join(' ') || 'powershell failed for an unknown reason'
  );
}

/** Encrypt a secret for this Windows user. */
export async function protectSecret(plain: string): Promise<string> {
  if (!canProtect()) {
    throw new Error(
      'Storing the mailbox password securely is only supported on Windows. ' +
        'On this machine, set IMAP_PASS in the environment instead of putting it in config.json.',
    );
  }
  const result = await runPowerShell(PROTECT_SCRIPT, plain);
  if (result.code !== 0 || !result.stdout.trim()) {
    throw new Error(`could not protect the password: ${cleanError(result)}`);
  }
  return result.stdout.trim();
}

/** Decrypt a blob made by `protectSecret` on this machine, by this user. */
export async function revealSecret(blob: string): Promise<string> {
  if (!canProtect()) {
    throw new Error(
      'config.json holds a Windows-protected password, which cannot be read on this system. ' +
        'Set IMAP_PASS in the environment, or run `sheetflow set-password` on the Windows machine.',
    );
  }
  const result = await runPowerShell(REVEAL_SCRIPT, blob);
  if (result.code !== 0) {
    throw new Error(
      `could not read the stored mailbox password: ${cleanError(result)}\n` +
        'It is encrypted for one Windows user on one machine - re-run `sheetflow set-password` here.',
    );
  }
  return result.stdout;
}

/**
 * Ask for the password at the terminal and return the blob. The password is typed into PowerShell's
 * own secure prompt, so it never appears in this process's memory, arguments, or the config file.
 *
 * A secure prompt needs a real console. Reading from a pipe or a shortcut would block forever, so
 * that case is refused before PowerShell is started.
 */
export async function promptForSecret(): Promise<string> {
  if (!canProtect()) {
    throw new Error('`set-password` is only supported on Windows.');
  }
  if (!process.stdin.isTTY) {
    throw new Error(
      'set-password needs a console to type into - it is not reading a password from a pipe.\n' +
        'Open a Command Prompt or PowerShell window in this folder and run `sheetflow set-password`.',
    );
  }
  const result = await runPowerShell(PROMPT_SCRIPT, '', true);
  if (result.code !== 0 || !result.stdout.trim()) {
    throw new Error(cleanError(result));
  }
  return result.stdout.trim();
}
