// E9 Original recovery: for every corpus document, through the MCP server, is the exact original recoverable?
// Per document: hash the file on disk, open it (this creates the history and, from 0.5.0, stores the original),
// export the original to a new path and compare SHA-256, edit one paragraph and save, export again and compare,
// then restore the original in place and compare the working file. Servers without the tools record "tool_unavailable".
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, readJsonl, CORPUS, manifest, fullOutline } from '../lib/exp.mjs';

const out = jsonl('e9_original');
const dir = freshDir('e9');
const s = await connect(dir);
const sha = (f) => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
const tools = (await s.client.listTools()).tools.map((t) => t.name);
const hasTools = tools.includes('export_original_document') && tools.includes('restore_original_document');
for (const d of manifest()) {
  const file = path.join(dir, d.file);
  fs.copyFileSync(path.join(CORPUS, d.file), file);
  const rec = { file: d.file, kind: d.kind, bytes: d.bytes, sha_before: sha(file), tools_available: hasTools };
  try {
    const o = await s.call('open_document', { path: file });
    if (!o.json?.ok) { rec.error = 'open: ' + o.text.slice(0, 200); out(rec); continue; }
    const sid = o.json.sessionId;
    rec.original_available = o.json.original?.available ?? null;
    rec.original_reported_sha = o.json.original?.sha256 ?? null;
    if (!hasTools) { rec.result = 'tool_unavailable'; out(rec); await s.call('close_document', { sessionId: sid }); continue; }
    // 1. export before any edit
    const e1 = path.join(dir, `${d.file}.export1.docx`);
    const x1 = await s.call('export_original_document', { sessionId: sid, outPath: e1 });
    rec.export_before_edit_ok = !!x1.json?.ok && fs.existsSync(e1) && sha(e1) === rec.sha_before;
    if (!x1.json?.ok) rec.export_before_edit_error = x1.text.slice(0, 200);
    // 2. edit and save
    const rows = await fullOutline(s, sid);
    const p = rows.find((r) => /^[0-9A-F]{8}$/i.test(r.id) && r.text.length > 20) ?? rows.find((r) => /^[0-9A-F]{8}$/i.test(r.id));
    if (!p) { rec.error = 'no editable paragraph'; out(rec); await s.call('close_document', { sessionId: sid }); continue; }
    const b = await s.call('execute_batch', { sessionId: sid, operations: [{ action: 'replace', target: p.id, text: 'E9 edit ' + d.file }] });
    const sv = await s.call('save_document', { sessionId: sid, message: 'e9 edit' });
    rec.edit_ok = !!b.json?.ok; rec.save_ok = !!sv.json?.ok;
    rec.sha_after_edit = sha(file);
    rec.file_changed_by_edit = rec.sha_after_edit !== rec.sha_before;
    // 3. export after edit
    const e2 = path.join(dir, `${d.file}.export2.docx`);
    const x2 = await s.call('export_original_document', { sessionId: sid, outPath: e2 });
    rec.export_after_edit_ok = !!x2.json?.ok && fs.existsSync(e2) && sha(e2) === rec.sha_before;
    if (!x2.json?.ok) rec.export_after_edit_error = x2.text.slice(0, 200);
    // 4. restore in place
    const r = await s.call('restore_original_document', { sessionId: sid });
    rec.restore_tool_ok = !!r.json?.ok;
    if (!r.json?.ok) rec.restore_error = r.text.slice(0, 200);
    rec.sha_after_restore = sha(file);
    rec.restore_ok = rec.sha_after_restore === rec.sha_before;
    // 5. history shows the restore as a commit
    const h = await s.call('get_document_history', { sessionId: sid });
    rec.history_commits = h.json?.totalCommits ?? null;
    await s.call('close_document', { sessionId: sid });
  } catch (e) { rec.error = String(e).slice(0, 300); }
  out(rec);
  fs.rmSync(file, { force: true });
}
const rows = readJsonl('e9_original');
console.log('documents', rows.length, 'export-before-edit ok', rows.filter((r) => r.export_before_edit_ok).length, 'export-after-edit ok', rows.filter((r) => r.export_after_edit_ok).length, 'restore ok', rows.filter((r) => r.restore_ok).length, 'errors', rows.filter((r) => r.error).length);
await s.close();
