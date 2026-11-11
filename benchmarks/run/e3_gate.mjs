// E3: schema gate efficacy. Phase A: injected XML/package faults + valid controls. Phase B: raw-XML and bad-parameter edits.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl, ROOT, CORPUS, manifest } from '../lib/exp.mjs';

const sha = (f) => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
const out = jsonl('e3_gate');
const dir = freshDir('e3');
const fdir = path.join(ROOT, 'corpus', 'faults');
const fm = JSON.parse(fs.readFileSync(path.join(fdir, 'faults.json'), 'utf8'));
const s = await connect(dir);
const TRIVIAL = [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'bench trailing paragraph' }];

async function probe(label, srcFile, meta) {
  const rec = { phase: 'A', label, ...meta };
  const work = path.join(dir, `${label}.docx`);
  fs.copyFileSync(srcFile, work);
  rec.sha_before = sha(work);
  const o = await s.call('open_document', { path: work });
  rec.open = { ok: !!o.json?.ok, code: o.json?.code, message: o.json?.ok ? undefined : (o.json?.message || o.text).slice(0, 200) };
  if (!o.json?.ok) { out(rec); return; }
  const sid = o.json.sessionId;
  const v = await s.call('validate_document', { sessionId: sid });
  rec.validate = { ok: !!v.json?.ok, valid: v.json?.validation?.valid, errorCount: v.json?.validation?.errorCount, first: (v.json?.validation?.errors || []).slice(0, 2).map((e) => (e.description || '').slice(0, 140)), code: v.json?.code,
    pending_normalization: v.json?.pendingNormalization ?? v.json?.pendingNormalizations ?? null };
  // strict save after a trivial edit
  const b = await s.call('execute_batch', { sessionId: sid, operations: TRIVIAL });
  rec.edit = { ok: !!b.json?.ok, code: b.json?.code };
  if (b.json?.ok) {
    const sv = await s.call('save_document', { sessionId: sid, strict: true, message: 'strict' });
    rec.strict = { ok: !!sv.json?.ok, code: sv.json?.code, normalized: sv.json?.normalized?.length ?? null, message: sv.json?.ok ? undefined : (sv.json?.message || '').slice(0, 160) };
    rec.strict.file_changed = sha(work) !== rec.sha_before;
  }
  await s.call('close_document', { sessionId: sid });
  // default-mode save on a fresh copy
  const work2 = path.join(dir, `${label}__default.docx`);
  fs.copyFileSync(srcFile, work2);
  const o2 = await s.call('open_document', { path: work2 });
  if (o2.json?.ok) {
    const b2 = await s.call('execute_batch', { sessionId: o2.json.sessionId, operations: TRIVIAL });
    if (b2.json?.ok) {
      const sv2 = await s.call('save_document', { sessionId: o2.json.sessionId, message: 'default' });
      rec.default = { ok: !!sv2.json?.ok, code: sv2.json?.code, normalized: sv2.json?.normalized?.length ?? null, valid: sv2.json?.validation?.valid, errorCount: sv2.json?.validation?.errorCount, message: sv2.json?.ok ? undefined : (sv2.json?.message || '').slice(0, 160) };
      rec.default.file_changed = sha(work2) !== rec.sha_before;
    }
    await s.call('close_document', { sessionId: o2.json.sessionId });
  }
  out(rec);
}

// clean corpus documents (false alarm baseline)
for (const d of manifest()) await probe(`clean__${d.file.replace('.docx', '')}`, path.join(CORPUS, d.file), { group: 'clean', base: d.file });
for (const f of fm.faults.filter((x) => x.applied)) await probe(f.file.replace('.docx', ''), path.join(fdir, f.file), { group: 'fault', id: f.id, category: f.category, kind: f.kind, title: f.title, rule: f.rule, base: f.base });
for (const c of fm.controls.filter((x) => x.applied)) await probe(c.file.replace('.docx', ''), path.join(fdir, c.file), { group: 'control', id: c.id, title: c.title, base: c.base });
console.log('phase A done');

// ---------- Phase B: edits that try to introduce invalid content ----------
const BASE = path.join(CORPUS, 'pd-mixed-010p.docx');
const PNG_BAD = Buffer.from('this is not an image at all').toString('base64');
const XML = {
  'x01 jc=middle': '<w:p><w:pPr><w:jc w:val="middle"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
  'x02 unknown element': '<w:p><w:bogus/></w:p>',
  'x03 sz=abc': '<w:p><w:r><w:rPr><w:sz w:val="abc"/></w:rPr><w:t>x</w:t></w:r></w:p>',
  'x04 table without rows': '<w:tbl><w:tblPr/><w:tblGrid/></w:tbl>',
  'x05 malformed xml': '<w:p><w:r><w:t>unclosed</w:r></w:p>',
  'x06 empty drawing': '<w:p><w:r><w:drawing/></w:r></w:p>',
  'x07 bookmark without id': '<w:p><w:bookmarkStart w:name="n"/></w:p>',
  'x08 pPr order': '<w:p><w:pPr><w:jc w:val="center"/><w:pStyle w:val="Heading1"/></w:pPr></w:p>',
  'x09 color=ZZZZZZ': '<w:p><w:r><w:rPr><w:color w:val="ZZZZZZ"/></w:rPr><w:t>x</w:t></w:r></w:p>',
  'x10 tc without p': '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc></w:tc></w:tr></w:tbl>',
  'x11 ins without author': '<w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p>',
  'x12 foreign namespace': '<w:p xmlns:foo="http://example.org"><foo:bar/></w:p>',
  'x13 xml:space=bad': '<w:p><w:r><w:t xml:space="bad">x</w:t></w:r></w:p>',
  'x14 text outside p': '<w:r><w:t>run directly in body</w:t></w:r>',
  'x15 VALID paragraph': '<w:p><w:r><w:t>valid raw xml paragraph</w:t></w:r></w:p>',
  'x16 VALID centred': '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>valid centred</w:t></w:r></w:p>',
};
const TYPED = {
  't01 color=banana': (id) => ({ action: 'style', target: id, color: 'banana' }),
  't02 fontSize=-1': (id) => ({ action: 'style', target: id, fontSize: -1 }),
  't03 fontSize=0': (id) => ({ action: 'style', target: id, fontSize: 0 }),
  't04 fontSize=99999': (id) => ({ action: 'style', target: id, fontSize: 99999 }),
  't05 alignment=middle': (id) => ({ action: 'style', target: id, alignment: 'middle' }),
  't06 highlight=purple': (id) => ({ action: 'style', target: id, highlight: 'purple' }),
  't07 mergeCells out of range': (id, tid) => ({ action: 'table', target: tid, operation: 'mergeCells', fromRow: 99, fromCol: 1, toRow: 99, toCol: 2 }),
  't08 image not base64': (id) => ({ action: 'image', operation: 'insert', anchor: 'end', fileData: '!!!not base64!!!' }),
  't09 image bytes not an image': (id) => ({ action: 'image', operation: 'insert', anchor: 'end', fileData: PNG_BAD }),
  't10 chart non-numeric values': (id) => ({ action: 'chart', operation: 'insert', chartType: 'column', anchor: 'end', categories: ['a'], series: [{ name: 's', values: ['x'] }] }),
  't11 chart no series': (id) => ({ action: 'chart', operation: 'insert', chartType: 'column', anchor: 'end', categories: ['a'], series: [] }),
  't12 hyperlink javascript:': (id) => ({ action: 'hyperlink', operation: 'insert', target: id, text: 'x', url: 'javascript:alert(1)' }),
  't13 bookmark name with spaces': (id) => ({ action: 'bookmark', operation: 'insert', target: id, bookmarkName: 'a b c' }),
  't14 unknown named style': (id) => ({ action: 'namedStyle', operation: 'apply', target: id, styleName: 'NoSuchStyle' }),
  't15 NUL in text': (id) => ({ action: 'replace', target: id, text: 'bad\u0000text' }),
  't16 negative margin': (id) => ({ action: 'pageSetup', marginTop: -5 }),
  't17 columns=99': (id) => ({ action: 'pageSetup', columns: 99 }),
  't18 empty comment': (id) => ({ action: 'comment', operation: 'insert', target: id, text: '' }),
  't19 ragged table rows': (id) => ({ action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['a', 'b', 'c'], rows: [['1'], ['1', '2', '3', '4']] }),
  't20 VALID bold': (id) => ({ action: 'style', target: id, bold: true }),
};
async function phaseB(name, buildOps, group) {
  const work = path.join(dir, `B-${group}-${name.split(' ')[0]}.docx`);
  fs.copyFileSync(BASE, work);
  const rec = { phase: 'B', group, name };
  const o = await s.call('open_document', { path: work });
  const sid = o.json.sessionId;
  const ol = await s.call('inspect_document', { sessionId: sid, limit: 40 });
  const row = ol.json.toonText.split('\n').find((l) => /^[0-9A-F]{8} \|/.test(l) && l.split(' | ')[5]?.length > 40);
  const id = row?.split(' | ')[0];
  const tid = (ol.json.toonText.match(/^(t_\d+) \|/m) || [])[1];
  rec.sha_before = sha(work);
  const b = await s.call('execute_batch', { sessionId: sid, operations: buildOps(id, tid) });
  rec.batch = { ok: !!b.json?.ok, code: b.json?.code ?? (b.inputRejected ? 'TOOL_INPUT_SCHEMA' : undefined), message: b.json?.ok ? undefined : (b.json?.message || b.text).slice(0, 200) };
  if (b.json?.ok) {
    const sv = await s.call('save_document', { sessionId: sid, message: name });
    rec.save = { ok: !!sv.json?.ok, code: sv.json?.code, valid: sv.json?.validation?.valid, errorCount: sv.json?.validation?.errorCount, message: sv.json?.ok ? undefined : (sv.json?.message || '').slice(0, 200) };
    rec.written = sha(work) !== rec.sha_before;
  }
  rec.outcome = !b.json?.ok ? 'rejected-at-edit' : rec.save?.ok ? 'saved' : 'blocked-at-save';
  rec.file = work;
  await s.call('close_document', { sessionId: sid });
  out(rec); console.log('B', group, name.padEnd(30), rec.outcome, rec.batch.code || rec.save?.code || '');
}
for (const [name, xml] of Object.entries(XML)) await phaseB(name, () => [{ action: 'insert', type: 'xml', anchor: 'end', position: 'after', xml }], 'raw-xml');
for (const [name, fn] of Object.entries(TYPED)) await phaseB(name, (id, tid) => [fn(id, tid)], 'typed');
await s.close();

// independent re-check of everything that phase B wrote
const written = fs.readFileSync(path.join(ROOT, 'results', JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).dependencies['@attentionform/astra-mcp-server'], 'e3_gate.jsonl'), 'utf8').split('\n').filter(Boolean).map(JSON.parse).filter((r) => r.phase === 'B' && r.written).map((r) => r.file);
if (written.length) {
  const list = path.join(dir, 'b-written.txt'); fs.writeFileSync(list, written.join('\n'));
  const o = execFileSync('python3', [path.join(ROOT, 'lib/oracles.py'), '--workers', '6', '--list', list], { maxBuffer: 1 << 26 }).toString().split('\n').filter(Boolean).map(JSON.parse);
  const f = jsonl('e3_phaseB_oracles'); o.forEach(f);
  console.log('phase B written files', written.length, 'pass O1-O5', o.filter((x) => x.pass_O1_O5).length);
}
