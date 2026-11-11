// E1: edit correctness and locality. Per trial: A (original), N (no-op save), E (edit + save).
import fs from 'node:fs';
import path from 'node:path';
import { connect } from '../lib/mcp.mjs';
import { manifest, freshDir, jsonl, CORPUS, WORK, fullOutline, seededRng, strHash, diffDocsAsync } from '../lib/exp.mjs';

const EDITS = ['replace', 'style-word', 'insert', 'delete', 'table-cell', 'comment', 'tracked-replace', 'named-style'];
const POSITIONS = ['first', 'middle', 'last', 'random'];
const WORKERS = Number(process.env.E1_WORKERS || 6);
const only = process.env.E1_DOCS ? process.env.E1_DOCS.split(',') : null;
const out = jsonl('e1_edit_locality');
const base = freshDir('e1');
const t0 = Date.now();

function pick(list, pos, rng) {
  if (!list.length) return null;
  return pos === 'first' ? list[0] : pos === 'last' ? list.at(-1) : pos === 'middle' ? list[Math.floor(list.length / 2)] : list[Math.floor(rng() * list.length)];
}

function buildEdit(edit, T, tag) {
  const newText = `Benchmark edit ${tag} replaced this text`;
  switch (edit) {
    case 'replace': return { ops: [{ action: 'replace', target: T.id, text: newText }], mode: 'same', expect_text: newText };
    case 'style-word': { const w = T.text.split(' ').find((x) => /^[a-z]{5,}$/i.test(x)); return { ops: [{ action: 'style', target: T.id, match: w, bold: true }], mode: 'same', expect_text: T.text, expect_raw: ['<w:b'] }; }
    case 'insert': return { ops: [{ action: 'insert', type: 'paragraph', anchor: T.id, position: 'after', text: newText }], mode: 'insert', expect_text: newText };
    case 'delete': return { ops: [{ action: 'delete', target: T.id }], mode: 'delete' };
    case 'table-cell': return { ops: [{ action: 'replace', target: T.id, text: newText }], mode: 'same', expect_text: newText };
    case 'comment': return { ops: [{ action: 'comment', operation: 'insert', target: T.id, text: `Comment ${tag}` }], mode: 'same', expect_text: T.text };
    case 'tracked-replace': return { ops: [{ action: 'replace', target: T.id, text: newText, trackChanges: true, author: 'Bench' }], mode: 'same', expect_raw: ['<w:ins'] };
    case 'named-style': return { ops: [{ action: 'namedStyle', operation: 'apply', target: T.id, styleName: 'Heading2' }], mode: 'same', expect_text: T.text, expect_raw: ['Heading2'] };
  }
}

async function saveTrial(s, file, ops, message) {
  const rec = {};
  const o = await s.call('open_document', { path: file });
  rec.open = { ok: !!o.json?.ok, ms: o.ms, code: o.json?.code };
  if (!o.json?.ok) return rec;
  const sid = o.json.sessionId;
  if (ops) {
    const b = await s.call('execute_batch', { sessionId: sid, operations: ops });
    rec.batch = { ok: !!b.json?.ok, ms: b.ms, code: b.json?.code, message: b.json?.ok ? undefined : (b.json?.message || b.text).slice(0, 200) };
    if (!b.json?.ok) { await s.call('close_document', { sessionId: sid }); return rec; }
  }
  const sv = await s.call('save_document', { sessionId: sid, message });
  rec.save = { ok: !!sv.json?.ok, ms: sv.ms, code: sv.json?.code, valid: sv.json?.validation?.valid, errorCount: sv.json?.validation?.errorCount, message: sv.json?.ok ? undefined : (sv.json?.message || sv.text).slice(0, 200) };
  await s.call('close_document', { sessionId: sid });
  return rec;
}

async function worker(wi, docs) {
  const dir = path.join(base, `w${wi}`); fs.mkdirSync(dir, { recursive: true });
  const s = await connect(dir);
  for (const doc of docs) {
    const stem = doc.file.replace('.docx', '');
    const A = path.join(dir, `${stem}__A.docx`);
    fs.copyFileSync(path.join(CORPUS, doc.file), A);
    // outline from the original (read-only session)
    const op = await s.call('open_document', { path: A });
    if (!op.json?.ok) { out({ doc: doc.file, stage: 'outline-open', error: op.text.slice(0, 200) }); continue; }
    const rows = await fullOutline(s, op.json.sessionId);
    await s.call('close_document', { sessionId: op.json.sessionId });
    const counts = new Map();
    for (const r of rows) counts.set(r.text, (counts.get(r.text) || 0) + 1);
    const paras = rows.filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length >= 60 && counts.get(r.text) === 1 && !r.text.startsWith('[IMG'));
    const tables = rows.filter((r) => /^t_\d+$/.test(r.id));
    const cells = [];
    for (const t of tables) {
      const m = [...t.text.matchAll(/r3c2 ([^|;]*)/g)][0];
      const txt = m?.[1]?.trim();
      if (txt && !t.text.includes('(span') ) cells.push({ id: `${t.id}:r3c2`, text: txt });
    }
    const cellCounts = new Map(); for (const r of rows) for (const m of r.text.matchAll(/r\d+c\d+ ([^|;]*)/g)) cellCounts.set(m[1].trim(), (cellCounts.get(m[1].trim()) || 0) + 1);
    const uniqCells = cells.filter((c) => cellCounts.get(c.text) === 1 && !paras.some((p) => p.text === c.text));

    // N: no-op save
    const Nf = path.join(dir, `${stem}__N.docx`); fs.copyFileSync(path.join(CORPUS, doc.file), Nf);
    const nrec = await saveTrial(s, Nf, null, 'noop');
    out({ doc: doc.file, kind: 'noop', ...nrec, N: Nf.replace(base, '') });
    if (!nrec.save?.ok) continue;
    const ndiff = await diffDocsAsync(A, Nf);
    out({ doc: doc.file, kind: 'first-save-normalization', diff: ndiff });

    for (const edit of EDITS) {
      for (const pos of POSITIONS) {
        const rng = seededRng(strHash(doc.file + edit));
        const T = pick(edit === 'table-cell' ? uniqCells : paras, pos, rng);
        const rec = { doc: doc.file, writer: doc.writer, pages: doc.measured_pages, kind: 'edit', edit, pos };
        if (!T) { out({ ...rec, skipped: 'no suitable target' }); continue; }
        const e = buildEdit(edit, T, `${edit}-${pos}`);
        const Ef = path.join(dir, `${stem}__${edit}__${pos}.docx`); fs.copyFileSync(path.join(CORPUS, doc.file), Ef);
        const tr = await saveTrial(s, Ef, e.ops, `bench ${edit} ${pos}`);
        Object.assign(rec, tr, { target: T.id, E: Ef.replace(base, '') });
        if (tr.save?.ok) {
          const spec = { target_text: T.text, mode: e.mode, expect_text: e.expect_text, expect_raw: e.expect_raw };
          if (!spec.expect_text) delete spec.expect_text; if (!spec.expect_raw) delete spec.expect_raw;
          const d = await diffDocsAsync(Nf, Ef, spec);
          rec.diff = { parts_changed: d.parts_changed, parts_changed_semantic: d.parts_changed_semantic, parts_added: d.parts_added, parts_removed: d.parts_removed, file_bytes_before: d.file_bytes_before, file_bytes_after: d.file_bytes_after, locality: d.locality, paragraphs_content_changed_before: d.paragraphs_content_changed_before, paragraphs_content_changed_after: d.paragraphs_content_changed_after };
        }
        out(rec);
      }
    }
    console.log(`[${((Date.now() - t0) / 1000).toFixed(0)}s] w${wi} done ${doc.file}`);
  }
  await s.close();
}

const docs = manifest().filter((d) => !only || only.includes(d.file)).sort((a, b) => b.bytes - a.bytes);
const buckets = Array.from({ length: WORKERS }, () => []);
docs.forEach((d, i) => buckets[i % WORKERS].push(d));
await Promise.all(buckets.map((b, i) => worker(i, b)));
console.log('E1 done', ((Date.now() - t0) / 1000).toFixed(0) + 's');
