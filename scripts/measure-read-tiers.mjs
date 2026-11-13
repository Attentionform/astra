#!/usr/bin/env node
/**
 * Sizes of the inspect tiers on a document this script generates, so anyone can repeat the numbers in
 * docs/reading-documents.md (part 2). Needs the CLI built (`npm run build`).
 *
 *   node scripts/measure-read-tiers.mjs [--chapters <n>] [--keep <dir>]
 *
 * The document is deterministic: <n> chapters (default 12), each with a Heading 1, three Heading 2 sections of
 * two Heading 3 parts, body paragraphs, and one table. It prints, per tier, the rows, the characters of the
 * output and an estimate of tokens (characters / 3.8, the estimate `scripts/measure-toon.ts` and the server use;
 * it is not a tokenizer). --keep writes the document and the raw outputs into <dir>.
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const cli = join(root, 'packages/cli/dist/cli.js');
const args = process.argv.slice(2);
const flag = (name, dflt) => (args.includes(name) ? args[args.indexOf(name) + 1] : dflt);
const chapters = Number(flag('--chapters', '12'));
const keep = flag('--keep', undefined);

const words = 'the quarterly plan covers staffing budget risk schedule vendor scope review owner target metric milestone delivery cost revenue'.split(' ');
let seed = 7;
const next = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff);
const sentence = () => `${Array.from({ length: 14 }, () => words[next() % words.length]).join(' ')}.`;
const paragraph = () => Array.from({ length: 4 }, sentence).join(' ');

let md = '';
for (let c = 1; c <= chapters; c++) {
  md += `# Chapter ${c}\n\n${paragraph()}\n\n`;
  for (let s = 1; s <= 3; s++) {
    md += `## Section ${c}.${s}\n\n${paragraph()}\n\n`;
    for (let p = 1; p <= 2; p++) md += `### Part ${c}.${s}.${p}\n\n${paragraph()}\n\n${paragraph()}\n\n`;
  }
  md += '| Item | Owner | Due |\n| --- | --- | --- |\n| One | A | Q1 |\n| Two | B | Q2 |\n\n';
}

const dir = keep ? (mkdirSync(keep, { recursive: true }), resolve(keep)) : mkdtempSync(join(tmpdir(), 'astra-tiers-'));
const mdPath = join(dir, 'tiers.md');
const docx = join(dir, 'tiers.docx');
writeFileSync(mdPath, md);
execFileSync('node', [cli, 'create', docx, '--from-md', mdPath, '--allowed-dir', dir, '--force'], { stdio: 'pipe' });

const run = (extra) => execFileSync('node', [cli, '--json', 'inspect', docx, ...extra], { encoding: 'utf8', maxBuffer: 1 << 28 });
const rows = [];
for (const tier of ['outline', 'compact', 'standard']) {
  const json = JSON.parse(run(['--fidelity', tier]).trim().split('\n').filter((l) => l.startsWith('{')).pop());
  if (keep) writeFileSync(join(dir, `${tier}.txt`), json.toon);
  rows.push({ tier, rows: json.rowCount, characters: json.toon.length, 'est. tokens': Math.round(json.toon.length / 3.8) });
}
const stats = JSON.parse(execFileSync('node', [cli, '--json', 'stats', docx], { encoding: 'utf8' }).trim().split('\n').filter((l) => l.startsWith('{')).pop());
process.stdout.write(`document: ${chapters} chapters, ${stats.words} words, ${stats.tables} tables\n`);
console.table(rows);
if (keep) process.stdout.write(`kept in ${dir}\n`);
