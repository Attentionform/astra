#!/usr/bin/env node
/**
 * Word parity check: renders each file with Microsoft Word and with LibreOffice and compares them.
 * Real programs, no mocks. Run after `npm run build`.
 *
 *   node scripts/word-parity.mjs [--out <dir>] [--strict] [--expect-word] [--dpi <n>] [file.docx|file.md ...]
 *
 * With no file the parity set in tests/fixtures/parity is used. A .md file is first turned into a .docx with `astra create --from-md`.
 * For every file the script writes <name>.word.pdf, <name>.lo.pdf and one PNG per page for each renderer, and prints
 *   PARITY <name> word=<pages|unavailable> lo=<pages> <same|DIFFERENT>
 * Word runs only on Windows with Word 2010 or later; elsewhere it prints `word=unavailable` (a failure only with --expect-word).
 * A different page count is reported; with --strict it is a failure. Look at the PNGs (Word first) before each release:
 * pagination equal is necessary, not sufficient (fonts, spacing, table widths, chart sizes).
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const cli = path.join(root, 'packages', 'cli', 'dist', 'cli.js');
const args = process.argv.slice(2);
const flag = (name) => args.includes(name);
const value = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? args[i + 1] : fallback;
};
const out = path.resolve(value('--out', fs.mkdtempSync(path.join(os.tmpdir(), 'astra-parity-'))));
const dpi = value('--dpi', '80');
const strict = flag('--strict');
const expectWord = flag('--expect-word');
const valued = new Set(['--out', '--dpi']);
const files = [];
for (let i = 0; i < args.length; i++) {
  if (valued.has(args[i])) { i++; continue; }
  if (!args[i].startsWith('--')) files.push(path.resolve(args[i]));
}
if (files.length === 0) {
  const dir = path.join(root, 'tests', 'fixtures', 'parity');
  for (const f of fs.readdirSync(dir).sort()) if (/\.(md|docx)$/.test(f)) files.push(path.join(dir, f));
}
fs.mkdirSync(out, { recursive: true });

function run(cmd, cmdArgs) {
  return spawnSync(cmd, cmdArgs, { encoding: 'utf8' });
}
function pages(pdf) {
  const r = run('pdfinfo', [pdf]);
  const m = /^Pages:\s+(\d+)/m.exec(r.stdout || '');
  return m ? Number(m[1]) : null;
}
function render(docx, renderer, target) {
  const r = run(process.execPath, [cli, 'export', docx, '--format', 'pdf', '--renderer', renderer, '-o', target, '--force']);
  return r.status === 0 && fs.existsSync(target) ? { ok: true } : { ok: false, why: `${r.stdout || ''}${r.stderr || ''}`.trim().split('\n').slice(-2).join(' ') };
}

let failed = false;
for (const file of files) {
  const name = path.basename(file).replace(/\.[^.]+$/, '');
  let docx = file;
  if (file.endsWith('.md')) {
    docx = path.join(out, `${name}.docx`);
    const made = run(process.execPath, [cli, 'create', docx, '--from-md', file, '--allowed-dir', path.dirname(file), '--force']);
    if (made.status !== 0) { console.log(`PARITY ${name} could not be created: ${(made.stderr || made.stdout).trim()}`); failed = true; continue; }
  }
  const result = {};
  for (const [key, renderer] of [['word', 'word'], ['lo', 'libreoffice']]) {
    const pdf = path.join(out, `${name}.${key}.pdf`);
    const done = render(docx, renderer, pdf);
    if (!done.ok) { result[key] = { why: done.why }; continue; }
    result[key] = { pages: pages(pdf) };
    run('pdftoppm', ['-r', dpi, '-png', pdf, path.join(out, `${name}.${key}`)]);
  }
  const w = result.word.pages;
  const l = result.lo.pages;
  const verdict = w === undefined ? 'word-unavailable' : w === l ? 'same' : 'DIFFERENT';
  console.log(`PARITY ${name} word=${w ?? 'unavailable'} lo=${l ?? 'failed'} ${verdict}`);
  if (w === undefined && result.word.why) console.log(`  word: ${result.word.why}`);
  if (l === undefined) { console.log(`  libreoffice: ${result.lo.why}`); failed = true; }
  if (w === undefined && expectWord) failed = true;
  if (verdict === 'DIFFERENT' && strict) failed = true;
}
console.log(`Output: ${out}`);
process.exit(failed ? 1 : 0);
