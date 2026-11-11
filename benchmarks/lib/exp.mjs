import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import AdmZip from 'adm-zip';
import { execFileSync } from 'node:child_process';

const here = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.resolve(here, '..');
export const RELEASE = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).dependencies['@attentionform/astra-mcp-server'];
export const RESULTS = path.join(ROOT, 'results', RELEASE);
export const WORK = path.join(ROOT, 'work');
export const CORPUS = path.join(ROOT, 'corpus', 'files');
fs.mkdirSync(RESULTS, { recursive: true });
fs.mkdirSync(WORK, { recursive: true });

export const manifest = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'corpus', 'manifest.json'), 'utf8')).documents;
export function freshDir(name) {
  const d = path.join(WORK, name);
  fs.rmSync(d, { recursive: true, force: true });
  fs.mkdirSync(d, { recursive: true });
  return d;
}
export function jsonl(name) {
  const f = path.join(RESULTS, name + '.jsonl');
  fs.writeFileSync(f, '');
  return (obj) => fs.appendFileSync(f, JSON.stringify(obj) + '\n');
}
export const readJsonl = (name) =>
  fs.readFileSync(path.join(RESULTS, name + '.jsonl'), 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l));
export const writeJson = (name, obj) => fs.writeFileSync(path.join(RESULTS, name + '.json'), JSON.stringify(obj, null, 1));
export const part = (file, name) => new AdmZip(file).readFile(name)?.toString('utf8');
export const diffDocs = (a, b) => JSON.parse(execFileSync('python3', [path.join(here, 'docdiff.py'), a, b], { maxBuffer: 1 << 28 }));
export function oracles(files, { lo = true, workers = 10 } = {}) {
  if (!files.length) return [];
  const list = path.join(WORK, `oracle-list-${process.pid}.txt`);
  fs.writeFileSync(list, files.join('\n'));
  const out = execFileSync('python3', [path.join(here, 'oracles.py'), '--workers', String(workers), ...(lo ? [] : ['--no-lo']), '--list', list], { maxBuffer: 1 << 29 }).toString();
  return out.split('\n').filter(Boolean).map((l) => JSON.parse(l));
}
export function rssKb(pid) {
  try { return Number(/VmRSS:\s+(\d+)/.exec(fs.readFileSync(`/proc/${pid}/status`, 'utf8'))[1]); } catch { return NaN; }
}

import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
const execFileP = promisify(execFile);
export const diffDocsAsync = async (a, b, spec) =>
  JSON.parse((await execFileP('python3', [path.join(here, 'docdiff.py'), a, b, ...(spec ? [JSON.stringify(spec)] : [])], { maxBuffer: 1 << 28 })).stdout);

// Parse the TOON outline text into rows.
export function parseToon(text) {
  const rows = [];
  for (const line of text.split('\n')) {
    if (!line || line.startsWith('>') || line.startsWith('ID |') || line.startsWith('---')) continue;
    const parts = line.split(' | ');
    if (parts.length < 6) continue;
    const [id, v, p, header, style, ...rest] = parts;
    const flagsAndText = rest;
    rows.push({ id: id.trim(), style: style.trim(), text: flagsAndText.slice(0, -1).join(' | ').trim(), flags: flagsAndText.at(-1).trim() });
  }
  return rows;
}
export function seededRng(seed) {
  let a = seed >>> 0;
  return () => { a = (a + 0x6d2b79f5) | 0; let t = Math.imul(a ^ (a >>> 15), 1 | a); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}
export const strHash = (s) => { let h = 2166136261; for (const c of s) { h ^= c.charCodeAt(0); h = Math.imul(h, 16777619); } return h >>> 0; };
export async function fullOutline(s, sessionId) {
  let text = '', offset = 0, hasMore = true;
  while (hasMore) {
    const r = await s.call('inspect_document', { sessionId, fidelity: 'standard', maxTextLength: 10000, offset, limit: 200 });
    if (!r.json?.ok) throw new Error('inspect failed: ' + r.text.slice(0, 200));
    text += r.json.toonText + '\n';
    hasMore = r.json.hasMore; offset = r.json.nextOffset ?? offset + 200;
  }
  return parseToon(text);
}
