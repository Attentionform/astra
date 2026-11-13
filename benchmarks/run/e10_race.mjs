// E10 Two writers, one document: two MCP server processes (two agents) open the same file, each edits a different
// paragraph, both save at the same time. Counts how often one edit is lost, how often a save fails, whether the file
// is still a readable package, and what the history looks like. Documents the known one-writer-per-document limit.
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import AdmZip from 'adm-zip';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS } from '../lib/exp.mjs';

const out = jsonl('e10_race');
const TRIALS = Number(process.env.E10_TRIALS || 30);
const SRC = path.join(CORPUS, process.env.E10_DOC || 'pd-prose-005p.docx');
const W = freshDir('e10');
const A = await connect(W), B = await connect(W);
const ids = async (c, sid) => {
  const r = await c.call('inspect_document', { sessionId: sid, offset: 0, limit: 40 });
  return (r.json?.toonText ?? r.text).split('\n').map((l) => l.trim().split(/\s*\|\s*/)[0].split(/\s+/)[0]).filter((x) => /^[A-Za-z0-9_:-]{3,}$/.test(x));
};
for (let i = 0; i < TRIALS; i++) {
  const p = path.join(W, `t${i}.docx`); fs.copyFileSync(SRC, p);
  const rec = { trial: i };
  try {
    const [oa, ob] = await Promise.all([A.call('open_document', { path: p }), B.call('open_document', { path: p })]);
    rec.open_errors = [oa, ob].filter((o) => o.isError || o.json?.ok === false).length;
    if (rec.open_errors) { rec.open_error_sample = [oa, ob].find((o) => o.isError || o.json?.ok === false).text.slice(0, 200); out(rec); continue; }
    const ia = await ids(A, oa.json.sessionId), ib = await ids(B, ob.json.sessionId);
    const ta = ia.find((x) => /^[0-9A-F]{8}$/i.test(x)) ?? ia[1];
    const tb = ib.filter((x) => /^[0-9A-F]{8}$/i.test(x))[2] ?? ib[3];
    const [ba, bb] = await Promise.all([
      A.call('execute_batch', { sessionId: oa.json.sessionId, operations: [{ action: 'replace', target: ta, text: `AGENTA${i} edit` }] }),
      B.call('execute_batch', { sessionId: ob.json.sessionId, operations: [{ action: 'replace', target: tb, text: `AGENTB${i} edit` }] }),
    ]);
    const [sa, sb] = await Promise.all([A.call('save_document', { sessionId: oa.json.sessionId, message: 'agent A' }), B.call('save_document', { sessionId: ob.json.sessionId, message: 'agent B' })]);
    rec.batch_ok = [!!ba.json?.ok, !!bb.json?.ok];
    rec.save_ok_flags = [!!sa.json?.ok, !!sb.json?.ok];
    rec.save_errors = [sa, sb].filter((x) => x.isError || x.json?.ok === false).length;
    const se = [sa, sb].find((x) => x.isError || x.json?.ok === false); if (se) rec.save_error_sample = se.text.slice(0, 200);
    await A.call('close_document', { sessionId: oa.json.sessionId }); await B.call('close_document', { sessionId: ob.json.sessionId });
    let xml = ''; try { xml = new AdmZip(p).readAsText('word/document.xml'); rec.readable = true; } catch { rec.readable = false; }
    const a = xml.includes(`AGENTA${i} edit`), b = xml.includes(`AGENTB${i} edit`);
    rec.edit_a_survives = a; rec.edit_b_survives = b; rec.both_survive = a && b; rec.one_lost = !!xml && (a !== b); rec.both_lost = !!xml && !a && !b;
    const repo = path.join(W, '.astra', `t${i}.docx`);
    try { rec.commits = Number(execFileSync('git', ['-C', repo, 'rev-list', '--count', 'HEAD'], { stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim()); } catch { rec.commits = null; }
    try { execFileSync('git', ['-C', repo, 'fsck', '--no-dangling'], { stdio: 'ignore' }); rec.fsck_ok = true; } catch { rec.fsck_ok = false; }
  } catch (e) { rec.error = String(e).slice(0, 300); }
  out(rec);
}
await A.close(); await B.close();
