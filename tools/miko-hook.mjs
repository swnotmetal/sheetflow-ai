#!/usr/bin/env node
// The Miko hook, with this host's tool names translated to the ones koma-miko looks for.
//
// `koma-miko`'s VS Code profile recognises `read_file` / `readFile` and `run_in_terminal`; this agent
// host reports `Read`, `view`, `Bash`, `powershell`, `skill`. Without the translation the gate can
// never observe a Skill being loaded, so `sheetflow-write-safety` denies (or, in review mode, pauses)
// every edit to src/excel.ts and friends no matter what the agent actually read.
//
// Only names are renamed. Nothing is added, dropped or invented: an edit still arrives as an edit and
// still faces the contract. A read of `…/SKILL.md` still has to actually happen for the evidence to
// be recorded — this file cannot manufacture it.
import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const TOOL_ALIASES = {
  Read: 'read_file',
  view: 'read_file',
  read: 'read_file',
  Bash: 'run_in_terminal',
  powershell: 'run_in_terminal',
  shell: 'run_in_terminal',
};

const hookCli = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  'node_modules',
  'koma-miko',
  'dist',
  'vscode-hook-cli.js',
);

let input;
try {
  input = JSON.parse(readFileSync(0, 'utf8'));
} catch {
  // Miko is fail-open on its own operational failures; an unreadable event is one of them.
  process.exit(0);
}

const alias = input && typeof input.tool_name === 'string' ? TOOL_ALIASES[input.tool_name] : undefined;
if (alias) input.tool_name = alias;

const result = spawnSync(process.execPath, [hookCli], {
  input: JSON.stringify(input),
  encoding: 'utf8',
});
if (result.stdout) process.stdout.write(result.stdout);
