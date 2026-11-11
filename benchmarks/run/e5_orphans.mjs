// E5 follow-up: a file reverted past a comment keeps two unreferenced parts. Does it still pass the independent checks?
import fs from 'node:fs';
import path from 'node:path';
import { validateFile } from '@attentionform/ooxml-validator';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, CORPUS, fullOutline, oracles } from '../lib/exp.mjs';

const out = jsonl('e5_orphans');
const dir = freshDir('e5-orphans');
const s = await connect(dir);
const DOCS = ['pd-prose-005p', 'pd-mixed-010p', 'pd-tables-005p', 'pd-lists-005p', 'pd-tracked-005p', 'lo-prose-005p', 'lo-mixed-010p', 'pd-comments-005p'];
const files = [];
for (const base of DOCS) {
  const f = path.join(dir, base + '.docx'); fs.copyFileSync(path.join(CORPUS, base + '.docx'), f);
  const o = await s.call('open_document', { path: f }); const sid = o.json.sessionId;
  const c0 = await s.call('save_document', { sessionId: sid, message: 'c0' });
  const baseline = await validateFile(f);
  const rows = (await fullOutline(s, sid)).filter((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 30);
  await s.call('execute_batch', { sessionId: sid, operations: [{ action: 'comment', operation: 'insert', target: rows[0].id, text: 'review me' }] });
  await s.call('save_document', { sessionId: sid, message: 'c1 comment' });
  await s.call('revert_document', { sessionId: sid, commitOid: c0.json.commitOid });
  const after = await validateFile(f);
  out({ base, errors_before_comment: baseline.errors?.length ?? 0, errors_after_revert: after.errors?.length ?? 0, file: f });
  files.push(f);
  await s.call('close_document', { sessionId: sid });
}
await s.close();
const o = oracles(files, { lo: true, workers: 6 });
const f2 = jsonl('e5_orphans_oracles'); o.forEach(f2);
console.log('reverted-past-comment files', files.length, 'pass O1-O5', o.filter((x) => x.pass_O1_O5).length);
