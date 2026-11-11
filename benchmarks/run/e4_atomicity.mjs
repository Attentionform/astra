// E4: a batch with one failing operation must leave the session exactly as it was.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS, fullOutline, seededRng } from '../lib/exp.mjs';

const out = jsonl('e4_atomicity');
const dir = freshDir('e4');
const s = await connect(dir);
const SIZES = [2, 5, 20, 100];
const POS = ['first', 'middle', 'last'];
const KINDS = {
  'unknown-target': () => ({ action: 'replace', target: 'DEADBEEF', text: 'x' }),
  'schema-invalid-param': (id) => ({ action: 'style', target: id, alignment: 'middle' }),
  'unknown-style': (id) => ({ action: 'namedStyle', operation: 'apply', target: id, styleName: 'NoSuchStyle' }),
  'absent-match': (id) => ({ action: 'style', target: id, match: 'zzzzqqqq', bold: true }),
};
const REPS = 3;
const hash = (t) => crypto.createHash('sha256').update(t).digest('hex');
let n = 0;
for (let rep = 0; rep < REPS; rep++) {
  for (const k of SIZES) for (const pos of POS) for (const [kind, mk] of Object.entries(KINDS)) {
    const name = `a${n++}`;
    const file = path.join(dir, name + '.docx');
    fs.copyFileSync(path.join(CORPUS, 'pd-prose-100p.docx'), file);
    const o = await s.call('open_document', { path: file });
    const sid = o.json.sessionId;
    const rows = await fullOutline(s, sid);
    const paras = rows.filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 40);
    const rng = seededRng(rep * 1000 + k * 10 + POS.indexOf(pos));
    const picks = [];
    while (picks.length < k) { const p = paras[Math.floor(rng() * paras.length)]; if (!picks.includes(p)) picks.push(p); }
    const failAt = pos === 'first' ? 0 : pos === 'last' ? k - 1 : Math.floor(k / 2);
    const ops = picks.map((p, i) => i === failAt ? KINDS[kind](p.id) : { action: 'replace', target: p.id, text: `atomic trial ${name} op ${i}` });
    const snap = async () => {
      const i = await s.call('inspect_document', { sessionId: sid, fidelity: 'standard', maxTextLength: 10000, limit: 200 });
      const rows2 = await fullOutline(s, sid);
      const h = await s.call('get_document_history', { sessionId: sid });
      return { outline: hash(rows2.map((r) => `${r.id}|${r.style}|${r.text}`).join('\n')), revision: i.json.revision, dirty: i.json.dirty, commits: h.json?.totalCommits };
    };
    const before = await snap();
    const b = await s.call('execute_batch', { sessionId: sid, operations: ops });
    const after = await snap();
    const rec = { name, rep, batch_size: k, fail_pos: pos, kind, batch_ok: !!b.json?.ok, code: b.json?.code, before, after,
      atomic: !b.json?.ok && before.outline === after.outline && before.revision === after.revision && before.dirty === after.dirty && before.commits === after.commits };
    // session still usable: the same batch without the failing op must apply
    const good = ops.filter((_, i) => i !== failAt);
    const g = await s.call('execute_batch', { sessionId: sid, operations: good });
    rec.recovery_ok = !!g.json?.ok && g.json.appliedCount === good.length;
    await s.call('close_document', { sessionId: sid, discard: true });
    out(rec);
    fs.rmSync(file, { force: true });
  }
}
const all = fs.readFileSync(path.join(path.dirname(new URL(import.meta.url).pathname), '..', 'results', '0.4.1', 'e4_atomicity.jsonl'), 'utf8').split('\n').filter(Boolean).map(JSON.parse);
console.log('trials', all.length, 'atomic', all.filter((r) => r.atomic).length, 'recovered', all.filter((r) => r.recovery_ok).length, 'batch accepted despite bad op', all.filter((r) => r.batch_ok).length);
await s.close();
