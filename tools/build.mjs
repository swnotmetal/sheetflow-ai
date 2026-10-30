// Single-executable build: bundle to CommonJS, then embed it into a copy of node.exe.
//
// The one thing that cannot go inside the executable is a native module: `@libsql/…` is a compiled
// `.node` binary, and `require()` has to find it on disk. So it is copied beside the executable into
// a `node_modules/` folder — no compiler, no `npm install`, just files that travel together.
//
//   node tools/build.mjs            → dist/sheetflow/  (exe + native module + config)
import { build } from 'esbuild';
import { copyFile, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

const OUT = 'dist/sheetflow';
const BUNDLE = 'dist/cli.cjs';

const platformKey = `${process.platform}-${process.arch}`;
const libsqlPkg = {
  'win32-x64': '@libsql/win32-x64-msvc',
  'darwin-x64': '@libsql/darwin-x64',
  'darwin-arm64': '@libsql/darwin-arm64',
  'linux-x64': '@libsql/linux-x64-gnu',
  'linux-arm64': '@libsql/linux-arm64-gnu',
}[platformKey];

if (!libsqlPkg) throw new Error(`no libsql build known for ${platformKey}`);

console.log(`bundling for ${platformKey}…`);
await build({
  entryPoints: ['src/cli.ts'],
  bundle: true,
  platform: 'node',
  // CommonJS on purpose: a single-executable application cannot be an ES module.
  format: 'cjs',
  target: 'node22',
  outfile: BUNDLE,
  external: ['node:*', '@libsql/*'],
  logLevel: 'warning',
  // Inside a single executable, `require` only resolves built-in modules — every package is meant to
  // be bundled in. `@libsql/*` cannot be: its `.node` binary is compiled code that has to sit on
  // disk. So a failed lookup falls back to a resolver anchored at the executable's own folder, where
  // the build places `node_modules/@libsql/…`. Nothing else changes: a lookup that would have failed
  // still fails, just from a place that also knows about the files next to the exe.
  banner: {
    js: [
      'const __nodeRequire = require;',
      "const { createRequire: __createRequire } = __nodeRequire('node:module');",
      'const __exeRequire = __createRequire(process.execPath);',
      'require = (id) => {',
      '  try { return __nodeRequire(id); }',
      '  catch (error) {',
      "    if (error && error.code === 'ERR_UNKNOWN_BUILTIN_MODULE') return __exeRequire(id);",
      '    throw error;',
      '  }',
      '};',
    ].join('\n'),
  },
});

await rm(OUT, { recursive: true, force: true });
await mkdir(OUT, { recursive: true });

// 1. The native module, exactly where Node's resolver will look for it.
const nativeSrc = `node_modules/${libsqlPkg}`;
const nativeDst = join(OUT, 'node_modules', libsqlPkg);
await mkdir(nativeDst, { recursive: true });
for (const file of ['index.node', 'package.json']) {
  if (existsSync(join(nativeSrc, file))) await copyFile(join(nativeSrc, file), join(nativeDst, file));
}

// 2. The executable: node itself, with the bundle injected.
const seaConfig = {
  main: BUNDLE,
  output: 'dist/sea-prep.blob',
  disableExperimentalSEAWarning: true,
  useSnapshot: false,
  useCodeCache: false,
};
await writeFile('dist/sea-config.json', JSON.stringify(seaConfig, null, 2));
execFileSync(process.execPath, ['--experimental-sea-config', 'dist/sea-config.json'], {
  stdio: 'inherit',
});

const exe = join(OUT, process.platform === 'win32' ? 'sheetflow.exe' : 'sheetflow');
await copyFile(process.execPath, exe);

// The sentinel `postject` looks for, and the fuse that tells Node an embedded blob is present.
const blob = await readFile('dist/sea-prep.blob');
if (process.platform === 'win32') {
  execFileSync(
    'npx',
    ['--yes', 'postject@1.0.0-alpha.6', exe, 'NODE_SEA_BLOB', 'dist/sea-prep.blob',
      '--sentinel-fuse', 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2'],
    { stdio: 'inherit', shell: true },
  );
} else {
  execFileSync(
    'npx',
    ['--yes', 'postject@1.0.0-alpha.6', exe, 'NODE_SEA_BLOB', 'dist/sea-prep.blob',
      '--sentinel-fuse', 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2',
      '--macho-segment-name', 'NODE_SEA'],
    { stdio: 'inherit', shell: true },
  );
}

// 3. A config the owner can read and edit, in place of a `.env` full of variable names. Written only
//    when absent: rebuilding the executable must never silently replace the shop's own answers.
const configPath = join(OUT, 'config.json');
if (!existsSync(configPath)) {
  await writeFile(
    configPath,
    `${JSON.stringify(
      {
        workbook: 'orders.xlsx',
        llmProfile: 'openai',
        llmApiKey: '',
        imap: { host: '', port: 993, user: '', password: '', mailbox: 'INBOX', limit: 10 },
        // A folder of .eml files, so the tool can be tried before a mailbox is connected.
        mailDir: 'mail',
        reviewTimeoutMs: 600000,
      },
      null,
      2,
    )}\n`,
  );
}

console.log(`\n${OUT}/`);
console.log(`  sheetflow.exe   embedded Node ${process.version} + the bundled pipeline`);
console.log(`  config.json     workbook, mailbox and model`);
console.log(`  node_modules/   the one native module Node cannot embed`);
