import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import AdmZip from 'adm-zip';

const execFileAsync = promisify(execFile);
const here = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.resolve(here, '..', 'fixtures');
const GALLERY = path.resolve(here, '..', '..', '..', '..', 'test-gallery');

export interface CorpusDoc {
  name: string;
  bytes: Buffer;
}

export interface GoldenError {
  id: string | null;
  path: string | null;
  xPath: string | null;
  errorType: string | null;
  /** sha256 of the description; the full text can run to many kilobytes. */
  description: string;
}

export interface GoldenEntry {
  exit: number;
  /** false when the program printed something other than JSON (unreadable input). */
  json: boolean;
  ok?: boolean;
  errors?: GoldenError[];
}

type Mutation = (entry: string, data: Buffer) => Buffer | null;

const replaceIn =
  (target: string, from: string, to: string): Mutation =>
  (entry, data) =>
    entry === target ? Buffer.from(data.toString('utf8').replace(from, to), 'utf8') : data;

const MUTATIONS: Record<string, Mutation> = {
  badelem: replaceIn('word/document.xml', '<w:body>', '<w:body><w:bogus/>'),
  badattr: (entry, data) =>
    entry === 'word/document.xml'
      ? Buffer.from(data.toString('utf8').replace(/<w:jc w:val="[a-z]+"/, '<w:jc w:val="zzz"'), 'utf8')
      : data,
  badorder: replaceIn('word/document.xml', '<w:p>', '<w:tbl></w:tbl><w:p>'),
  malformed: replaceIn('word/document.xml', '</w:body>', ''),
  badstyles: replaceIn('word/styles.xml', '<w:styles', '<w:styles x:q="1"'),
  nocontenttypes: (entry, data) => (entry === '[Content_Types].xml' ? null : data),
  nomain: (entry, data) => (entry === 'word/document.xml' ? null : data),
};

function mutate(source: Buffer, fn: Mutation): Buffer {
  const zip = new AdmZip(source);
  const out = new AdmZip();
  for (const e of zip.getEntries()) {
    if (e.isDirectory) continue;
    const next = fn(e.entryName, e.getData());
    if (next) out.addFile(e.entryName, next);
  }
  return out.toBuffer();
}

/**
 * A fixed set of valid, invalid and unreadable documents: the package fixtures as they are, plus every
 * gallery document with each mutation applied, a truncated copy, and two non-zip inputs. Nothing random.
 */
export function buildCorpus(): CorpusDoc[] {
  const docs: CorpusDoc[] = [];
  for (const f of fs.readdirSync(FIXTURES).sort()) {
    docs.push({ name: `fixture__${f}`, bytes: fs.readFileSync(path.join(FIXTURES, f)) });
  }
  const gallery = fs.readdirSync(GALLERY).filter((f) => f.endsWith('.docx')).sort();
  for (const f of gallery) {
    const bytes = fs.readFileSync(path.join(GALLERY, f));
    const base = f.replace(/\.docx$/, '');
    docs.push({ name: `${base}.docx`, bytes });
    for (const [key, fn] of Object.entries(MUTATIONS)) {
      docs.push({ name: `${base}__${key}.docx`, bytes: mutate(bytes, fn) });
    }
    docs.push({ name: `${base}__truncated.docx`, bytes: bytes.subarray(0, Math.floor(bytes.length / 2)) });
  }
  docs.push({ name: 'garbage.docx', bytes: Buffer.from('not a zip at all') });
  docs.push({ name: 'empty.docx', bytes: Buffer.alloc(0) });
  return docs;
}

const sha256 = (s: string): string => createHash('sha256').update(s).digest('hex');

/** Runs a validator binary once on a document (one-shot mode, default Office version). */
export async function runOneShot(binary: string, file: string): Promise<GoldenEntry> {
  let exit = 0;
  let stdout = '';
  try {
    ({ stdout } = await execFileAsync(binary, [file], { maxBuffer: 256 * 1024 * 1024, timeout: 60_000 }));
  } catch (err) {
    const e = err as { code?: number; stdout?: string };
    exit = typeof e.code === 'number' ? e.code : 1;
    stdout = e.stdout ?? '';
  }
  let parsed: { ok?: boolean; errors?: Record<string, unknown>[] };
  try {
    parsed = JSON.parse(stdout);
  } catch {
    return { exit, json: false };
  }
  const str = (v: unknown): string | null => (typeof v === 'string' ? v : null);
  return {
    exit,
    json: true,
    ok: parsed.ok,
    errors: (parsed.errors ?? []).map((e) => ({
      id: str(e.id),
      path: str(e.path),
      xPath: str(e.xPath),
      errorType: str(e.errorType),
      description: sha256(str(e.description) ?? ''),
    })),
  };
}

/** Writes the corpus to a fresh temp directory and returns it. The caller removes it. */
export function writeCorpus(docs: CorpusDoc[]): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ooxml-parity-'));
  for (const d of docs) fs.writeFileSync(path.join(dir, d.name), d.bytes);
  return dir;
}

export async function runCorpus(binary: string, docs: CorpusDoc[]): Promise<Record<string, GoldenEntry>> {
  const dir = writeCorpus(docs);
  try {
    const out: Record<string, GoldenEntry> = {};
    const queue = [...docs];
    const workers = Array.from({ length: 4 }, async () => {
      for (let d = queue.shift(); d; d = queue.shift()) out[d.name] = await runOneShot(binary, path.join(dir, d.name));
    });
    await Promise.all(workers);
    return Object.fromEntries(Object.entries(out).sort(([a], [b]) => (a < b ? -1 : 1)));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}
