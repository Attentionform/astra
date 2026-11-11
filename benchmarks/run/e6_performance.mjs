// E6 latency/memory by size, E7 batch scaling. Run on an otherwise idle machine.
import fs from 'node:fs';
import path from 'node:path';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS, fullOutline, rssKb } from '../lib/exp.mjs';

const out = jsonl('e6_performance');
const SIZES = ['pd-prose-001p', 'pd-prose-010p', 'pd-prose-050p', 'pd-prose-100p', 'pd-prose-200p'];
const WARM = 2, RUNS = 15;
for (const base of SIZES) {
  const dir = freshDir('e6-' + base);
  const s = await connect(dir);
  let peak = 0; const timer = setInterval(() => { const r = rssKb(s.pid); if (r > peak) peak = r; }, 50);
  const rssIdle = rssKb(s.pid);
  for (let run = -WARM - 1; run < RUNS; run++) {
    const phase = run === -WARM - 1 ? 'cold' : run < 0 ? 'warmup' : 'timed';
    const file = path.join(dir, `r${run + 10}.docx`);
    fs.copyFileSync(path.join(CORPUS, base + '.docx'), file);
    const rec = { base, phase, run };
    const o = await s.call('open_document', { path: file }); rec.open_ms = o.ms;
    if (!o.json?.ok) { rec.error = o.text.slice(0, 200); out(rec); continue; }
    const sid = o.json.sessionId;
    const i = await s.call('inspect_document', { sessionId: sid }); rec.inspect_ms = i.ms; rec.inspect_rows = i.json?.returnedNodes;
    const rows = (await fullOutline(s, sid)).filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 40);
    const word = rows[Math.floor(rows.length / 2)].text.split(' ')[2];
    const se = await s.call('search_content', { sessionId: sid, query: word }); rec.search_ms = se.ms;
    const r1 = await s.call('execute_batch', { sessionId: sid, operations: [{ action: 'replace', target: rows[0].id, text: 'perf single edit' }] }); rec.edit1_ms = r1.ms;
    const ops50 = rows.slice(1, 51).map((r, k) => ({ action: 'replace', target: r.id, text: `perf batch ${k}` }));
    const r50 = await s.call('execute_batch', { sessionId: sid, operations: ops50 }); rec.edit50_ms = r50.ms; rec.edit50_ok = !!r50.json?.ok;
    const v = await s.call('validate_document', { sessionId: sid }); rec.validate_ms = v.ms;
    const sv = await s.call('save_document', { sessionId: sid, message: 'perf' }); rec.save_ms = sv.ms; rec.save_ok = !!sv.json?.ok;
    rec.validator_share_of_save = rec.validate_ms / rec.save_ms;
    await s.call('close_document', { sessionId: sid, discard: true });
    rec.rss_kb = rssKb(s.pid);
    out(rec);
  }
  clearInterval(timer);
  out({ base, phase: 'memory', rss_idle_kb: rssIdle, rss_peak_kb: peak });
  console.log(base, 'peak RSS MB', (peak / 1024).toFixed(0));
  await s.close();
}

// E7 batch scaling on a 100-page document
const e7 = jsonl('e7_batch_scaling');
const dir = freshDir('e7');
const s = await connect(dir);
for (const k of [1, 10, 50, 100, 250, 500]) for (let run = 0; run < 10; run++) {
  const file = path.join(dir, `b${k}-${run}.docx`); fs.copyFileSync(path.join(CORPUS, 'pd-prose-100p.docx'), file);
  const o = await s.call('open_document', { path: file }); const sid = o.json.sessionId;
  const rows = (await fullOutline(s, sid)).filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 40);
  const ops = rows.slice(0, k).map((r, j) => ({ action: 'replace', target: r.id, text: `scale ${j}` }));
  const b = await s.call('execute_batch', { sessionId: sid, operations: ops });
  const sv = await s.call('save_document', { sessionId: sid, message: 'scale' });
  e7({ ops: k, run, available: rows.length, batch_ok: !!b.json?.ok, batch_ms: b.ms, save_ms: sv.ms, ms_per_op: b.ms / k });
  await s.call('close_document', { sessionId: sid, discard: true }); fs.rmSync(file, { force: true });
}
await s.close();
