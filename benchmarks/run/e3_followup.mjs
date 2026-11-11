// E3 follow-up: validate (a) every injected file and (b) every file ASTra wrote in default mode with the
// standalone upstream validator (no ASTra session pipeline), to separate "not detected" from "healed before the verdict".
import fs from 'node:fs';
import path from 'node:path';
import { validateFile } from '@attentionform/ooxml-validator';
import { jsonl, readJsonl, WORK, ROOT, CORPUS, manifest } from '../lib/exp.mjs';

const out = jsonl('e3_followup');
const rows = readJsonl('e3_gate').filter((r) => r.phase === 'A');
const dir = path.join(WORK, 'e3');
const fdir = path.join(ROOT, 'corpus', 'faults');
const count = async (f) => { try { const r = await validateFile(f); return { errors: r.errors?.length ?? (r.ok ? 0 : -1), ok: r.ok, first: (r.errors || []).slice(0, 1).map((e) => (e.description || String(e)).slice(0, 120)) }; } catch (e) { return { errors: -2, ok: false, first: [String(e).slice(0, 120)] }; } };
const base = {};
for (const d of manifest()) base[d.file] = await count(path.join(CORPUS, d.file));
for (const r of rows) {
  const src = r.group === 'clean' ? path.join(CORPUS, r.base) : path.join(fdir, `${r.label}.docx`);
  const rec = { label: r.label, group: r.group, id: r.id, base: r.base, kind: r.kind, raw: await count(src) };
  const dflt = path.join(dir, `${r.label}__default.docx`);
  if (r.default?.file_changed && fs.existsSync(dflt)) rec.default_output = await count(dflt);
  out(rec);
}
console.log('e3 follow-up done', rows.length);
