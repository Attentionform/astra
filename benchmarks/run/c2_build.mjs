// C2: build documents from an empty document with typed operations. 15 fixed recipes.
import fs from 'node:fs';
import path from 'node:path';
import AdmZip from 'adm-zip';
import { connect } from '../lib/mcp.mjs';
import { freshDir, jsonl } from '../lib/exp.mjs';

const out = jsonl('c2_build');
const dir = freshDir('c2');
const s = await connect(dir);
// 1x1 red PNG, inline
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==';

const P = (text, anchor = 'end') => ({ action: 'insert', type: 'paragraph', anchor, position: 'after', text });
// each recipe: async ({batch}) with batch(ops) -> {ok, ids[], code, message}; plus check(zip) -> string[] of failures
const RECIPES = {
  'styled-headings': { run: async ({ batch }) => { const r = await batch([P('Alpha heading'), P('Beta heading'), P('Gamma heading')]); if (!r.ok) return r; return batch([0, 1, 2].map((i) => ({ action: 'namedStyle', operation: 'apply', target: r.ids[i], styleName: `Heading${i + 1}` }))); },
    check: (z) => ['Heading1', 'Heading2', 'Heading3'].filter((h) => !z.doc.includes(`w:val="${h}"`)) },
  'paragraph-format': { run: async ({ batch }) => { const r = await batch([P('Formatted paragraph text')]); if (!r.ok) return r; return batch([{ action: 'style', target: r.ids[0], alignment: 'center', bold: true, italic: true, color: 'FF0000', fontSize: 28, spacingBefore: 240 }]); },
    check: (z) => [['w:jc w:val="center"', 'center'], ['<w:b', 'bold'], ['<w:i', 'italic'], ['FF0000', 'color']].filter(([n]) => !z.doc.includes(n)).map((x) => x[1]) },
  'bullet-list': { run: async ({ batch }) => { const r = await batch([P('Item one'), P('Item two'), P('Item three')]); if (!r.ok) return r; return batch([{ action: 'list', operation: 'create', target: r.ids, listType: 'bullet' }]); },
    check: (z) => (z.doc.match(/<w:numPr>/g) || []).length >= 3 ? [] : ['numPr x3'] },
  'numbered-list': { run: async ({ batch }) => { const r = await batch([P('Step one'), P('Step two'), P('Step three')]); if (!r.ok) return r; return batch([{ action: 'list', operation: 'create', target: r.ids, listType: 'numbered' }]); },
    check: (z) => (z.doc.match(/<w:numPr>/g) || []).length >= 3 && z.numbering.includes('decimal') ? [] : ['numbered numPr/decimal'] },
  'table-merged': { run: async ({ batch }) => { const r = await batch([{ action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['Region', 'Q1', 'Q2'], rows: [['North', '10', '12'], ['South', '8', '9'], ['East', '7', '11']] }]); if (!r.ok) return r; const tid = r.ids[0]; return batch([{ action: 'table', target: tid, operation: 'mergeCells', fromRow: 2, fromCol: 2, toRow: 2, toCol: 3 }]); },
    check: (z) => z.doc.includes('gridSpan') || z.doc.includes('vMerge') || z.doc.includes('hMerge') ? [] : ['merge marker'] },
  'table-styled': { run: async ({ batch }) => { const r = await batch([{ action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']], headerBold: true }]); if (!r.ok) return r; return batch([{ action: 'table', target: r.ids[0], operation: 'setShading', fill: 'DDDDDD' }, { action: 'table', target: r.ids[0], operation: 'setHeaderRepeat' }]); },
    check: (z) => z.doc.includes('DDDDDD') ? [] : ['shading'] },
  'inline-image': { run: async ({ batch }) => batch([{ action: 'image', operation: 'insert', anchor: 'end', fileData: PNG, width: 2000000, height: 2000000, altText: 'red pixel' }]),
    check: (z) => z.names.some((n) => n.startsWith('word/media/')) && z.doc.includes('<w:drawing') ? [] : ['media+drawing'] },
  callout: { run: async ({ batch }) => batch([{ action: 'callout', text: 'Important note for the reader', title: 'Note', theme: 'indigo', anchor: 'end', position: 'after' }]),
    check: (z) => z.doc.includes('Important note for the reader') ? [] : ['callout text'] },
  chart: { run: async ({ batch }) => batch([{ action: 'chart', operation: 'insert', chartType: 'column', title: 'Revenue', anchor: 'end', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Revenue', values: [10, 14, 18] }] }]),
    check: (z) => z.names.some((n) => /^word\/charts\/chart\d*\.xml$/.test(n)) ? [] : ['chart part'] },
  equation: { run: async ({ batch }) => batch([{ action: 'equation', operation: 'insert', latex: 'E = mc^2', anchor: 'end', position: 'after' }]),
    check: (z) => z.doc.includes('m:oMath') ? [] : ['oMath'] },
  footnote: { run: async ({ batch }) => { const r = await batch([P('A sentence that needs a footnote.')]); if (!r.ok) return r; return batch([{ action: 'footnote', operation: 'insert', anchor: r.ids[0], text: 'The footnote text.' }]); },
    check: (z) => z.names.includes('word/footnotes.xml') && z.all('word/footnotes.xml').includes('The footnote text.') ? [] : ['footnotes part'] },
  'header-footer': { run: async ({ batch }) => batch([{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'Benchmark header', alignment: 'center' }, { action: 'headerFooter', operation: 'set', targetType: 'footer', text: 'Page footer', alignment: 'right' }]),
    check: (z) => z.names.some((n) => /word\/header\d*\.xml/.test(n)) && z.names.some((n) => /word\/footer\d*\.xml/.test(n)) ? [] : ['header/footer parts'] },
  'page-setup-columns': { run: async ({ batch }) => batch([P('Body text for page setup'), { action: 'pageSetup', orientation: 'landscape', paperSize: 'A4', marginTop: 1000, marginBottom: 1000, marginLeft: 1200, marginRight: 1200, columns: 2 }]),
    check: (z) => z.doc.includes('w:orient="landscape"') && z.doc.includes('w:cols') ? [] : ['landscape+cols'] },
  'section-break': { run: async ({ batch }) => batch([P('Before the break'), { action: 'insert', type: 'sectionBreak', anchor: 'end', position: 'after', breakType: 'nextPage' }, P('After the break')]),
    check: (z) => (z.doc.match(/<w:sectPr/g) || []).length >= 2 ? [] : ['two sectPr'] },
  'hyperlink-comment': { run: async ({ batch }) => { const r = await batch([P('Visit the documentation page for details')]); if (!r.ok) return r; const a = await batch([{ action: 'hyperlink', operation: 'insert', target: r.ids[0], text: 'documentation page', url: 'https://example.org/docs' }]); if (!a.ok) return a; return batch([{ action: 'comment', operation: 'insert', target: r.ids[0], text: 'Please review this sentence' }]); },
    check: (z) => (z.doc.includes('w:hyperlink') ? [] : ['hyperlink']).concat(z.names.includes('word/comments.xml') ? [] : ['comments part']) },
};

for (const [name, rc] of Object.entries(RECIPES)) {
  const file = path.join(dir, `${name}.docx`);
  const rec = { recipe: name };
  const n = await s.call('new_document', { outPath: file, template: 'blank', force: true });
  rec.new = { ok: !!n.json?.ok, code: n.json?.code };
  if (!n.json?.ok) { out(rec); continue; }
  const sid = n.json.sessionId;
  let calls = 0, ms = 0;
  const batch = async (operations) => {
    const r = await s.call('execute_batch', { sessionId: sid, operations }); calls++; ms += r.ms;
    const ids = [];
    for (const x of r.json?.results || []) ids.push(...(x.createdNodeIds || []));
    return { ok: !!r.json?.ok, ids, code: r.json?.code, message: (r.json?.message || '').slice(0, 300) };
  };
  let res;
  try { res = await rc.run({ batch }); } catch (e) { res = { ok: false, code: 'HARNESS_EXCEPTION', message: String(e).slice(0, 200) }; }
  rec.batch = { ok: !!res.ok, code: res.code, message: res.message, calls, ms };
  if (res.ok) {
    const v = await s.call('validate_document', { sessionId: sid });
    rec.validate = { valid: v.json?.validation?.valid, errorCount: v.json?.validation?.errorCount };
    const sv = await s.call('save_document', { sessionId: sid, message: `build ${name}` });
    rec.save = { ok: !!sv.json?.ok, code: sv.json?.code, valid: sv.json?.validation?.valid, errorCount: sv.json?.validation?.errorCount, ms: sv.ms, message: sv.json?.ok ? undefined : (sv.json?.message || sv.text).slice(0, 300) };
    if (sv.json?.ok) {
      const zip = new AdmZip(file);
      const names = zip.getEntries().map((e) => e.entryName);
      const read = (n) => zip.readAsText(n);
      const zc = { names, doc: read('word/document.xml'), numbering: names.includes('word/numbering.xml') ? read('word/numbering.xml') : '', all: read };
      rec.feature_missing = rc.check(zc);
      rec.file = file;
    }
  }
  await s.call('close_document', { sessionId: sid });
  out(rec);
  console.log(name.padEnd(20), 'batch', rec.batch.ok, rec.batch.code || '', 'save', rec.save?.ok, 'missing', JSON.stringify(rec.feature_missing || ''), rec.batch.message?.slice(0, 120) || '');
}
await s.close();
