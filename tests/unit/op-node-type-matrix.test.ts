import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const T2 = path.resolve(__dirname, '../../test-gallery/T2_edit_target.docx');

async function buildBaseBytes(): Promise<Buffer> {
  const parser = new DocxParser(fs.readFileSync(T2));
  const ast = await parser.getAST();
  const first = ast.sections[0].nodes[0].id;
  const engine = new OperationEngine(ast, parser);
  const r = await engine.executeBatch({
    rollbackOnError: true,
    operations: [
      { action: 'insert', type: 'paragraph', anchor: first, position: 'after', text: 'list item A', id: 'li_a' },
      { action: 'insert', type: 'paragraph', anchor: 'li_a', position: 'after', text: 'list item B', id: 'li_b' },
      { action: 'list', operation: 'create', target: ['li_a', 'li_b'], listType: 'bullet' },
      { action: 'equation', operation: 'insert', anchor: 'li_b', position: 'after', latex: 'x^2', id: 'eq_1' },
      { action: 'insert', type: 'chart', anchor: 'eq_1', position: 'after', chartType: 'column', title: 'T',
        categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }], id: 'chart_1' },
      { action: 'image', operation: 'insert', anchor: 'chart_1', position: 'after',
        fileData: 'data:image/png;base64,' + PNG, width: 914400, height: 914400, altText: 'a', id: 'img_1' },
      { action: 'image', operation: 'insert', anchor: 'img_1', position: 'after',
        fileData: 'data:image/png;base64,' + PNG, width: 914400, height: 914400, altText: 'b', id: 'img_2' },
    ] as never,
  });
  if (!r.success) throw new Error('matrix setup failed: ' + JSON.stringify(r.errors));
  const mixed = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === 'img_2')!.xml as Element;
  const run = mixed.ownerDocument.createElement('w:r');
  const t = mixed.ownerDocument.createElement('w:t');
  t.setAttribute('xml:space', 'preserve');
  t.textContent = 'Text beside the picture. ';
  run.appendChild(t);
  const pPr = mixed.getElementsByTagName('w:pPr')[0];
  mixed.insertBefore(run, pPr ? pPr.nextSibling : mixed.firstChild);
  return parser.rezip();
}

const NODE_TYPES = ['paragraph', 'table', 'equation', 'chart', 'image', 'mixed'] as const;
function pickIds(ast: Awaited<ReturnType<DocxParser['getAST']>>) {
  const nodes = ast.sections.flatMap((s) => s.nodes);
  const isMixed = (x: (typeof nodes)[number]) => x.type === 'paragraph' && x.metadata?.hasImage === true;
  const byType: Record<string, string> = {};
  for (const t of NODE_TYPES) {
    const n = nodes.find((x) =>
      t === 'mixed' ? isMixed(x)
      : x.type === t && (t !== 'paragraph' || ((x.metadata?.text ?? '').length > 0 && !isMixed(x))));
    if (n) byType[t] = n.id;
  }
  const otherParagraph = nodes.filter((x) => x.type === 'paragraph' && !isMixed(x)).map((x) => x.id)[1];
  return { byType, otherParagraph };
}

type Case = { name: string; build: (id: string, other: string) => Record<string, unknown> };
const CASES: Case[] = [
  { name: 'replace',          build: (id) => ({ action: 'replace', target: id, text: 'x' }) },
  { name: 'delete',           build: (id) => ({ action: 'delete', target: id }) },
  { name: 'style',            build: (id) => ({ action: 'style', target: id, bold: true }) },
  { name: 'table.updateCell', build: (id) => ({ action: 'table', target: id, params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'x' } }) },
  { name: 'chart.update',     build: (id) => ({ action: 'chart', operation: 'update', target: id, title: 'x' }) },
  { name: 'equation.replace', build: (id) => ({ action: 'equation', operation: 'replace', target: id, latex: 'x' }) },
  { name: 'list.create',      build: (id) => ({ action: 'list', operation: 'create', target: [id], listType: 'bullet' }) },
  { name: 'move',             build: (id, other) => ({ action: 'move', target: id, anchor: other, position: 'after' }) },
  { name: 'clearFormat',      build: (id) => ({ action: 'clearFormat', target: id }) },
  { name: 'namedStyle.apply', build: (id) => ({ action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: id }) },
  { name: 'hyperlink',        build: (id) => ({ action: 'hyperlink', target: id, url: 'https://example.com' }) },
  { name: 'bookmark',         build: (id) => ({ action: 'bookmark', target: id, name: 'bm1' }) },
  { name: 'image.delete',     build: (id) => ({ action: 'image', operation: 'delete', target: id }) },
  { name: 'image.resize',     build: (id) => ({ action: 'image', operation: 'resize', target: id, width: 914400, height: 914400 }) },
  { name: 'image.setAltText', build: (id) => ({ action: 'image', operation: 'setAltText', target: id, altText: 'a' }) },
  { name: 'image.setCaption', build: (id) => ({ action: 'image', operation: 'setCaption', target: id, captionText: 'c' }) },
  { name: 'image.setWrap',    build: (id) => ({ action: 'image', operation: 'setWrap', target: id, wrapMode: 'square' }) },
  { name: 'image.reposition', build: (id) => ({ action: 'image', operation: 'reposition', target: id, wrapMode: 'square' }) },
];

const ALL = ['paragraph', 'table', 'equation', 'chart', 'image', 'mixed'];
const NOT_TABLE = ['paragraph', 'equation', 'chart', 'image', 'mixed'];
const EXPECTED_ACCEPT: Record<string, string[]> = {
  'replace': ['paragraph', 'mixed'], 'delete': ALL,
  'style': ALL, 'table.updateCell': ['table'], 'chart.update': ['chart'],
  'equation.replace': ['equation', 'paragraph'], 'list.create': NOT_TABLE,
  'move': ALL,
  'clearFormat': ALL, 'namedStyle.apply': NOT_TABLE, 'hyperlink': NOT_TABLE,
  'bookmark': ALL,
  'image.delete': ['image', 'mixed'], 'image.resize': ['image', 'mixed'], 'image.setAltText': ['image', 'mixed'],
  'image.setCaption': ['image', 'mixed'], 'image.setWrap': ['image', 'mixed'], 'image.reposition': ['image', 'mixed'],
};

const count = (b: Buffer, re: RegExp) => (new AdmZip(b).readAsText('word/document.xml').match(re) ?? []).length;

describe('operation x node kind matrix', () => {
  it('runs every combination', async () => {
    const baseBytes = await buildBaseBytes();
    expect(Object.keys(pickIds(await new DocxParser(baseBytes).getAST()).byType).sort()).toEqual([...NODE_TYPES].sort());
    const untyped: string[] = [], changedOnRefuse: string[] = [], unexpectedAccepts: string[] = [], objectLoss: string[] = [], invalidAfterAccept: string[] = [], unexpectedRefusals: string[] = [];
    for (const c of CASES) {
      for (const type of NODE_TYPES) {
        const parser = new DocxParser(baseBytes);
        const ast = await parser.getAST();
        const { byType, otherParagraph } = pickIds(ast);
        if (!byType[type]) { untyped.push('NOTFOUND ' + type); continue; }
        const before = parser.rezip();
        const engine = new OperationEngine(ast, parser);
        const res: any = await engine.executeBatch({ operations: [c.build(byType[type], otherParagraph)] as never, rollbackOnError: true });
        const key = c.name + ' on ' + type;
        if (res.success === false) {
          const first = String(res.errors?.[0]?.error ?? '');
          if (!/E_(VALIDATION|NOT_FOUND)/.test(first)) untyped.push(key + ' :: ' + first.slice(0, 120));
          if (Buffer.compare(parser.rezip(), before) !== 0) changedOnRefuse.push(key);
          if (EXPECTED_ACCEPT[c.name].includes(type)) unexpectedRefusals.push(key + ' :: ' + first.slice(0, 140));
        } else {
          if (!EXPECTED_ACCEPT[c.name].includes(type)) unexpectedAccepts.push(key);
          const after = parser.rezip();
          if (c.name !== 'delete' && c.name !== 'image.delete') {
            for (const [label, re] of [['drawing', /<w:drawing[ >]/g], ['oMath', /<m:oMath[ >]/g], ['commentRef', /<w:commentReference /g]] as const) {
              if (count(after, re) < count(before, re)) objectLoss.push(key + ' (' + label + ')');
            }
          }
          const v = await validateDocxBuffer(after);
          if (!v.ok) invalidAfterAccept.push(key + ' :: ' + JSON.stringify(v.errors).slice(0, 140));
        }
      }
    }
    // errors[0] is the operation's own error; the last entry is the batch wrapper and always carries E_VALIDATION.
    expect({ untyped, changedOnRefuse, unexpectedAccepts, unexpectedRefusals, objectLoss, invalidAfterAccept }).toEqual({
      untyped: [], changedOnRefuse: [], unexpectedAccepts: [], unexpectedRefusals: [], objectLoss: [], invalidAfterAccept: [],
    });
  }, 180000);
});
