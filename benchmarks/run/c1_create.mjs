// C1: create documents from Markdown (create pillar) + templates + out-of-scope fail-closed checks.
import fs from 'node:fs';
import path from 'node:path';
import { encode } from 'gpt-tokenizer';
import { execFileSync } from 'node:child_process';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, part, ROOT } from '../lib/exp.mjs';
import { genMarkdown } from '../lib/mdgen.mjs';

const out = jsonl('c1_create');
const dir = freshDir('c1');
const s = await connect(dir);
const SIZES = [1, 1, 2, 2, 3, 5, 5, 8, 10, 10, 15, 20, 25, 25, 30, 40, 50, 50, 60, 75, 100, 100, 2, 5];
const trial = async (name, markdown, blocks, pages, seed, template = 'blank') => {
  const file = path.join(dir, name + '.docx');
  const rec = { name, kind: 'compose', template, pages_target: pages, seed, md_bytes: Buffer.byteLength(markdown), md_tokens: encode(markdown).length };
  const n = await s.call('new_document', { outPath: file, template, force: true });
  rec.new = { ok: !!n.json?.ok, ms: n.ms, code: n.json?.code };
  if (!n.json?.ok) return out(rec);
  const sid = n.json.sessionId;
  const c = await s.call('compose_document', { sessionId: sid, markdown, ...(template !== 'blank' ? { template } : {}) });
  rec.compose = { ok: !!c.json?.ok, ms: c.ms, code: c.json?.code, message: c.json?.ok ? undefined : (c.json?.message || c.text).slice(0, 300) };
  if (!c.json?.ok) { await s.call('close_document', { sessionId: sid, discard: true }); return out(rec); }
  const v = await s.call('validate_document', { sessionId: sid });
  rec.validate = { ok: !!v.json?.ok, valid: v.json?.validation?.valid, errorCount: v.json?.validation?.errorCount, ms: v.ms };
  const sv = await s.call('save_document', { sessionId: sid, message: `create ${name}` });
  rec.save = { ok: !!sv.json?.ok, ms: sv.ms, code: sv.json?.code, valid: sv.json?.validation?.valid, errorCount: sv.json?.validation?.errorCount, message: sv.json?.ok ? undefined : (sv.json?.message || sv.text).slice(0, 300) };
  await s.call('close_document', { sessionId: sid });
  if (sv.json?.ok) {
    const xml = part(file, 'word/document.xml');
    rec.out_xml_bytes = Buffer.byteLength(xml); rec.out_xml_tokens = encode(xml).length;
    rec.write_cost_ratio = rec.out_xml_tokens / rec.md_tokens;
    rec.file = file; rec.file_bytes = fs.statSync(file).size;
    const bj = path.join(dir, name + '.blocks.json'); fs.writeFileSync(bj, JSON.stringify(blocks));
    try { rec.fidelity = JSON.parse(execFileSync('python3', [path.join(ROOT, 'lib/fidelity.py'), file, bj]).toString()); }
    catch (e) { rec.fidelity = { error: String(e).slice(0, 200) }; }
  }
  out(rec);
  console.log(name, 'compose', rec.compose.ok, rec.compose.ms?.toFixed(0) + 'ms', 'save', rec.save?.ok, 'fid', rec.fidelity ? `${rec.fidelity.found_in_order}/${rec.fidelity.expected}` : '-', 'ratio', rec.write_cost_ratio?.toFixed(1));
};

SIZES.forEach(() => {});
for (let i = 0; i < SIZES.length; i++) {
  const g = genMarkdown(SIZES[i], 70000 + i);
  await trial(`md-${String(i).padStart(2, '0')}-${SIZES[i]}p`, g.markdown, g.blocks, SIZES[i], 70000 + i);
}
// templates: same small spec on each template
for (const template of ['report', 'memo']) {
  const g = genMarkdown(2, 71000);
  await trial(`tpl-${template}`, g.markdown, g.blocks, 2, 71000, template);
}

// Out-of-scope inputs must fail closed (E_VALIDATION, line number, nothing written).
const OOS = {
  'heading-level-4': '# T\n\n#### Too deep\n\ntext\n',
  'nested-list': '- a\n  - b\n- c\n',
  'task-list': '- [ ] todo\n- [x] done\n',
  'raw-html': '# T\n\n<div class="x">raw</div>\n',
  'nested-table-in-quote': '> | a | b |\n> | --- | --- |\n> | 1 | 2 |\n',
  'heading-in-quote': '> ## inside\n',
  'remote-image': '![x](https://example.org/a.png)\n',
  'image-without-allowedDir': '![x](pic.png)\n',
  'table-without-separator': '| a | b |\n| 1 | 2 |\n',
  'html-comment-script': '<script>alert(1)</script>\n',
  'deep-nested-quote': '> > > > four\n',
  'custom-style-block': '::: custom\ntext\n:::\n',
};
for (const [name, md] of Object.entries(OOS)) {
  const file = path.join(dir, `oos-${name}.docx`);
  const n = await s.call('new_document', { outPath: file, template: 'blank', force: true });
  const sid = n.json.sessionId;
  const before = (await s.call('inspect_document', { sessionId: sid })).json?.rowCount;
  const c = await s.call('compose_document', { sessionId: sid, markdown: md });
  const after = (await s.call('inspect_document', { sessionId: sid })).json?.rowCount;
  const rec = { name, kind: 'out-of-scope', accepted: !!c.json?.ok, code: c.json?.code, message: (c.json?.message || '').slice(0, 200), rows_before: before, rows_after: after, unchanged: before === after };
  out(rec); console.log('oos', name, rec.accepted ? 'ACCEPTED' : rec.code, 'unchanged', rec.unchanged);
  await s.call('close_document', { sessionId: sid, discard: true });
}
await s.close();
