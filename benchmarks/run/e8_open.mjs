// E8 open latency: first open of a document that has no history yet, and reopen of one that already has it.
// Same server process per case (warmed once), N documents per case, each opened twice: the first open creates
// the history, the second reopens it. A 200-page document, with and without the validator wait, and the same
// document with two 10 MiB incompressible media parts (about 21 MB). Raw samples are written; nothing is dropped.
import fs from 'node:fs';
import path from 'node:path';
import AdmZip from 'adm-zip';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS } from '../lib/exp.mjs';

const out = jsonl('e8_open');
const N = Number(process.env.E8_N || 15);
const dir0 = freshDir('e8-src');
const base = path.join(CORPUS, 'pd-mixed-200p.docx');
const heavy = path.join(dir0, 'pd-mixed-200p-20mb.docx');
{
  // Deterministic, incompressible filler (xorshift32), so the size does not depend on the compressor.
  const filler = (seed) => { const b = Buffer.alloc(10 * 1024 * 1024); let x = seed >>> 0; for (let i = 0; i < b.length; i += 4) { x ^= x << 13; x >>>= 0; x ^= x >>> 17; x ^= x << 5; x >>>= 0; b.writeUInt32LE(x, i); } return b; };
  const zip = new AdmZip(base);
  zip.addFile('word/media/big1.png', filler(0x9e3779b9));
  zip.addFile('word/media/big2.png', filler(0x85ebca6b));
  zip.writeZip(heavy);
}
const CASES = [
  { label: '200p', src: base, args: {} },
  { label: '200p-validation-off', src: base, args: { validationTimeoutMs: 0 } },
  { label: '200p-20mb-validation-off', src: heavy, args: { validationTimeoutMs: 0 } },
];
for (const cs of CASES) {
  const dir = freshDir('e8-' + cs.label);
  const s = await connect(dir);
  const open = async (file) => {
    const o = await s.call('open_document', { path: file, ...cs.args });
    if (!o.json?.ok) throw new Error(`open failed: ${o.text.slice(0, 200)}`);
    await s.call('close_document', { sessionId: o.json.sessionId });
    return o.ms;
  };
  const warm = path.join(dir, 'warm.docx'); fs.copyFileSync(cs.src, warm); await open(warm);
  for (let i = 0; i < N; i++) {
    const file = path.join(dir, `d${i}.docx`); fs.copyFileSync(cs.src, file);
    const rec = { case: cs.label, run: i, bytes: fs.statSync(file).size };
    try { rec.first_open_ms = await open(file); rec.reopen_ms = await open(file); } catch (e) { rec.error = String(e).slice(0, 200); }
    out(rec);
  }
  console.log(cs.label, 'done');
  await s.close();
}
