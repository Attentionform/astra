// E5: history, preview, revert, redo. Byte equality against hashes recorded at each save.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import AdmZip from 'adm-zip';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS, fullOutline, seededRng, strHash } from '../lib/exp.mjs';

const out = jsonl('e5_history');
const dir = freshDir('e5');
const s = await connect(dir);
const sha = (b) => crypto.createHash('sha256').update(b).digest('hex');
const docXml = (f) => sha(new AdmZip(f).readFile('word/document.xml'));
const fileSha = (f) => sha(fs.readFileSync(f));
const DOCS = ['pd-prose-005p', 'pd-mixed-010p', 'lo-tables-005p', 'lo-prose-005p', 'pd-tables-005p', 'pd-lists-005p', 'pd-comments-005p', 'pd-tracked-005p', 'lo-lists-005p', 'lo-mixed-010p'];
const SEEDS = [1, 2, 3];
const N = 10;
const partsOf = (f) => { const z = new AdmZip(f); return Object.fromEntries(z.getEntries().filter((e) => !e.isDirectory).map((e) => [e.entryName, sha(e.getData())])); };
const changedParts = (a, b) => [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => a[k] !== b[k]).sort();

function mkEdit(kind, p, i) {
  const t = `history edit ${i} ${kind}`;
  if (kind === 'replace') return { action: 'replace', target: p.id, text: t };
  if (kind === 'insert') return { action: 'insert', type: 'paragraph', anchor: p.id, position: 'after', text: t };
  if (kind === 'comment') return { action: 'comment', operation: 'insert', target: p.id, text: t };
  if (kind === 'tracked') return { action: 'replace', target: p.id, text: t, trackChanges: true, author: 'Bench' };
  return { action: 'namedStyle', operation: 'apply', target: p.id, styleName: 'Heading2' };
}
for (const base of DOCS) for (const seed of SEEDS) {
  const name = `${base}-s${seed}`;
  const file = path.join(dir, name + '.docx');
  fs.copyFileSync(path.join(CORPUS, base + '.docx'), file);
  const rec = { name, base, seed, saves: [] };
  const rng = seededRng(strHash(name));
  const o = await s.call('open_document', { path: file });
  const sid = o.json.sessionId;
  // initial save establishes commit 0 (first-save normalization included)
  const s0 = await s.call('save_document', { sessionId: sid, message: 'c0 initial' });
  rec.saves.push({ i: 0, ok: !!s0.json?.ok, oid: s0.json?.commitOid, file_sha: fileSha(file), doc_sha: docXml(file), parts: partsOf(file), bytes: fs.statSync(file).size });
  for (let i = 1; i <= N; i++) {
    const rows = await fullOutline(s, sid);
    const paras = rows.filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 40);
    const p = paras[Math.floor(rng() * paras.length)];
    const kind = ['replace', 'insert', 'comment', 'tracked', 'named'][Math.floor(rng() * 5)];
    const b = await s.call('execute_batch', { sessionId: sid, operations: [mkEdit(kind, p, i)] });
    const sv = await s.call('save_document', { sessionId: sid, message: `c${i} ${kind}` });
    rec.saves.push({ i, kind, batch_ok: !!b.json?.ok, ok: !!sv.json?.ok, oid: sv.json?.commitOid, file_sha: fileSha(file), doc_sha: docXml(file), parts: partsOf(file), bytes: fs.statSync(file).size });
  }
  const h = await s.call('get_document_history', { sessionId: sid, limit: 100 });
  rec.history = { total: h.json?.totalCommits, listed: h.json?.history?.length, oids_match: (h.json?.history || []).slice().reverse().map((x) => x.oid).join() === rec.saves.map((x) => x.oid).join() };
  rec.reverts = [];
  const head = rec.saves.at(-1);
  for (let k = 1; k <= N; k++) {
    const target = rec.saves[N - k];
    const pre = partsOf(file);
    const pv = await s.call('preview_revert', { sessionId: sid, steps: k });
    const rv = await s.call('revert_document', { sessionId: sid, steps: k });
    const post = partsOf(file);
    const r = { k, preview_ok: !!pv.json?.ok, revert_ok: !!rv.json?.ok, code: rv.json?.code,
      doc_xml_equal: docXml(file) === target.doc_sha, file_equal: fileSha(file) === target.file_sha,
      all_parts_equal: JSON.stringify(post) === JSON.stringify(target.parts),
      preview_files: pv.json?.affectedFiles, actual_changed_parts: changedParts(pre, post) };
    r.preview_matches_actual = JSON.stringify([...(pv.json?.affectedFiles || [])].sort()) === JSON.stringify(r.actual_changed_parts);
    const rd = await s.call('redo_document', { sessionId: sid });
    r.redo_ok = !!rd.json?.ok; r.redo_doc_xml_equal = docXml(file) === head.doc_sha; r.redo_file_equal = fileSha(file) === head.file_sha;
    rec.reverts.push(r);
  }
  await s.call('close_document', { sessionId: sid, discard: true });
  out(rec);
  const ok = rec.reverts.filter((r) => r.all_parts_equal).length;
  console.log(name.padEnd(26), 'commits', rec.history.total, 'parts-equal', ok + '/' + N, 'file-equal', rec.reverts.filter((r) => r.file_equal).length, 'redo-file-equal', rec.reverts.filter((r) => r.redo_file_equal).length);
}

// growth of the history store
const growth = jsonl('e5_growth');
for (const base of ['pd-prose-005p', 'pd-prose-050p', 'pd-mixed-100p']) {
  const file = path.join(dir, `growth-${base}.docx`);
  fs.copyFileSync(path.join(CORPUS, base + '.docx'), file);
  const o = await s.call('open_document', { path: file }); const sid = o.json.sessionId;
  const du = (d) => { let t = 0; const w = (p) => { for (const e of fs.readdirSync(p, { withFileTypes: true })) { const f = path.join(p, e.name); e.isDirectory() ? w(f) : (t += fs.statSync(f).size); } }; if (fs.existsSync(d)) w(d); return t; };
  await s.call('save_document', { sessionId: sid, message: 'g0' });
  growth({ base, saves: 0, docx_bytes: fs.statSync(file).size, history_bytes: du(path.join(dir, '.astra', path.basename(file))) });
  const rows = await fullOutline(s, sid);
  const paras = rows.filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 40);
  for (let i = 1; i <= 50; i++) {
    await s.call('execute_batch', { sessionId: sid, operations: [{ action: 'replace', target: paras[i % paras.length].id, text: `growth edit ${i}` }] });
    await s.call('save_document', { sessionId: sid, message: `g${i}` });
    if (i % 10 === 0) growth({ base, saves: i, docx_bytes: fs.statSync(file).size, history_bytes: du(path.join(dir, '.astra', path.basename(file))) });
  }
  await s.call('close_document', { sessionId: sid, discard: true });
}
await s.close();
