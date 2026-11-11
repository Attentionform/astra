// E5: history, preview, revert, redo. Byte equality against hashes recorded at each save.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import AdmZip from 'adm-zip';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS, fullOutline, seededRng, strHash } from '../lib/exp.mjs';

const WI = Number(process.env.E5_WORKER ?? 0), WN = Number(process.env.E5_WORKERS ?? 1);
const out = jsonl(WN > 1 ? `e5_history_w${WI}` : 'e5_history');
const dir = freshDir(WN > 1 ? `e5-w${WI}` : 'e5');
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
async function build(base, seed, tag) {
  const name = `${base}-s${seed}${tag}`;
  const file = path.join(dir, name + '.docx');
  fs.copyFileSync(path.join(CORPUS, base + '.docx'), file);
  const rng = seededRng(strHash(`${base}-s${seed}`));
  const o = await s.call('open_document', { path: file });
  const sid = o.json.sessionId;
  const saves = [];
  const s0 = await s.call('save_document', { sessionId: sid, message: 'c0 initial' });
  saves.push({ i: 0, ok: !!s0.json?.ok, oid: s0.json?.commitOid, file_sha: fileSha(file), doc_sha: docXml(file), parts: partsOf(file), bytes: fs.statSync(file).size });
  for (let i = 1; i <= N; i++) {
    const rows = await fullOutline(s, sid);
    const ids = rows.filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && !r.text.startsWith('[IMG') && r.text.length > 0);
    const long = ids.filter((r) => r.text.length > 40);
    const paras = long.length ? long : ids;
    const p = paras[Math.floor(rng() * paras.length)];
    const kind = ['replace', 'insert', 'comment', 'tracked', 'named'][Math.floor(rng() * 5)];
    const b = await s.call('execute_batch', { sessionId: sid, operations: [mkEdit(kind, p, i)] });
    const sv = await s.call('save_document', { sessionId: sid, message: `c${i} ${kind}` });
    saves.push({ i, kind, batch_ok: !!b.json?.ok, ok: !!sv.json?.ok, oid: sv.json?.commitOid, file_sha: fileSha(file), doc_sha: docXml(file), parts: partsOf(file), bytes: fs.statSync(file).size });
  }
  return { name, file, sid, saves };
}
// Parts reachable from the package root through relationships (what an application actually reads).
function reachable(file) {
  const z = new AdmZip(file); const names = new Set(z.getEntries().map((e) => e.entryName));
  const seen = new Set(['[Content_Types].xml']); const queue = [''];
  while (queue.length) {
    const src = queue.pop();
    const relName = src === '' ? '_rels/.rels' : `${path.posix.dirname(src)}/_rels/${path.posix.basename(src)}.rels`.replace(/^\.\//, '');
    if (!names.has(relName)) continue;
    seen.add(relName);
    for (const m of z.readAsText(relName).matchAll(/<Relationship\b[^>]*>/g)) {
      if (/TargetMode="External"/.test(m[0])) continue;
      const t = /Target="([^"]*)"/.exec(m[0])?.[1]; if (!t) continue;
      const full = t.startsWith('/') ? t.slice(1) : path.posix.normalize(path.posix.join(src === '' ? '' : path.posix.dirname(src), t));
      if (names.has(full) && !seen.has(full)) { seen.add(full); queue.push(full); }
    }
  }
  return seen;
}
const cmp = (file, target) => {
  const now = partsOf(file); const want = target.parts;
  const keys = [...new Set([...Object.keys(now), ...Object.keys(want)])];
  const reach = reachable(file);
  const diff = keys.filter((k) => now[k] !== want[k]);
  return { doc_xml_equal: docXml(file) === target.doc_sha, file_equal: fileSha(file) === target.file_sha, all_parts_equal: diff.length === 0,
    extra_parts: diff.filter((k) => !want[k]), missing_parts: diff.filter((k) => !now[k]), differing_parts: diff.filter((k) => now[k] && want[k]),
    reachable_equal: keys.filter((k) => reach.has(k)).every((k) => now[k] === want[k]) };
};

let seqIdx = 0;
for (const base of DOCS) for (const seed of SEEDS) {
  if (seqIdx++ % WN !== WI) continue;
  // (a) revert by commit id to every earlier saved state, each followed by redo
  const { name, file, sid, saves } = await build(base, seed, '');
  const rec = { name, base, seed, saves: saves.map(({ parts, ...x }) => x), reverts: [], steps: [] };
  const h = await s.call('get_document_history', { sessionId: sid, limit: 100 });
  const listed = (h.json?.history || []).map((x) => x.oid);
  rec.history = { total: h.json?.totalCommits, listed: listed.length, saves_all_listed: saves.every((x) => listed.includes(x.oid)), listed_order_matches: listed.slice(0, saves.length).reverse().join() === saves.map((x) => x.oid).join() };
  const head = saves.at(-1);
  for (let i = 0; i < N; i++) {
    const target = saves[i];
    const pre = partsOf(file);
    const pv = await s.call('preview_revert', { sessionId: sid, commitOid: target.oid });
    const rv = await s.call('revert_document', { sessionId: sid, commitOid: target.oid });
    const post = partsOf(file);
    const r = { to: i, depth: N - i, preview_ok: !!pv.json?.ok, revert_ok: !!rv.json?.ok, code: rv.json?.code, ...cmp(file, target), preview_files: pv.json?.affectedFiles, actual_changed_parts: changedParts(pre, post) };
    r.preview_matches_actual = JSON.stringify([...(pv.json?.affectedFiles || [])].sort()) === JSON.stringify(r.actual_changed_parts);
    const rd = await s.call('redo_document', { sessionId: sid });
    r.redo_ok = !!rd.json?.ok; r.redo = cmp(file, head);
    rec.reverts.push(r);
  }
  await s.call('close_document', { sessionId: sid });
  // (b) revert by step count, one fresh history per depth (steps count from the current head)
  for (const k of (seed === 1 ? [1, 2, 5, 9] : [])) {
    const b = await build(base, seed, `-k${k}`);
    const target = b.saves[N - k];
    const pv = await s.call('preview_revert', { sessionId: b.sid, steps: k });
    const rv = await s.call('revert_document', { sessionId: b.sid, steps: k });
    rec.steps.push({ k, preview_ok: !!pv.json?.ok, revert_ok: !!rv.json?.ok, code: rv.json?.code, ...cmp(b.file, target), preview_target_message: pv.json?.targetMessage, expected_message: `c${N - k}` });
    await s.call('close_document', { sessionId: b.sid });
  }
  out(rec);
  console.log(name.padEnd(26), 'reachable-equal', rec.reverts.filter((r) => r.reachable_equal).length, 'by-oid parts-equal', rec.reverts.filter((r) => r.all_parts_equal).length + '/' + N, 'file-equal', rec.reverts.filter((r) => r.file_equal).length, 'redo parts-equal', rec.reverts.filter((r) => r.redo.all_parts_equal).length, '| steps parts-equal', rec.steps.filter((r) => r.all_parts_equal).length + '/4');
}

// growth of the history store (worker 0 only)
if (WI !== 0) { await s.close(); process.exit(0); }
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
  await s.call('close_document', { sessionId: sid });
}
await s.close();
