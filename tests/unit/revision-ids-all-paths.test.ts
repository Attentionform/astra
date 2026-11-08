import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const TABLE = `<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>`
  + `<w:tr>${cell('alpha')}${cell('beta')}</w:tr><w:tr>${cell('gamma')}${cell('')}</w:tr></w:tbl>`;
const PARA = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello brave world</w:t></w:r></w:p>';

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const live = (parser: DocxParser, tag: 'ins' | 'del') => Array.from(parser.getDocumentDom().getElementsByTagName(`w:${tag}`)).map(e => Number(e.getAttribute('w:id')));
const idsInDocument = (parser: DocxParser) => [...live(parser, 'ins'), ...live(parser, 'del')].sort((a, b) => a - b);

describe('every tracked edit returns the ids of the revisions it created', () => {
  it('replace with the default path', async () => {
    const { parser, ast, engine } = await open(PARA);
    const r = await engine.executeBatch({ operations: [{ action: 'replace', target: ast.sections[0].nodes[0].id, text: 'hello brave new world', trackChanges: true }] } as never);
    expect(r.success).toBe(true);
    expect([...(r.revisionIds ?? [])].sort((a, b) => a - b)).toEqual(idsInDocument(parser));
    expect(r.revisionIds?.length).toBeGreaterThan(0);
  });

  for (const flags of [{ preserveFormatting: false }, { resetStyle: true }] as const) {
    it(`replace with ${JSON.stringify(flags)} (the non-preserve path)`, async () => {
      const { parser, ast, engine } = await open(PARA);
      const r = await engine.executeBatch({ operations: [{ action: 'replace', target: ast.sections[0].nodes[0].id, text: 'something else entirely', trackChanges: true, ...flags }] } as never);
      expect(r.success, JSON.stringify(r.errors)).toBe(true);
      const created = [...(r.revisionIds ?? [])].sort((a, b) => a - b);
      expect(created.length).toBeGreaterThan(0);
      expect(created).toEqual(idsInDocument(parser));
      // the ids are usable: rejecting them restores the original text
      const rej = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'reject', ids: created }] } as never);
      expect(rej.success, JSON.stringify(rej.errors)).toBe(true);
      expect(ast.sections[0].nodes[0].metadata.text).toBe('hello brave world');
    });
  }

  it('delete with trackChanges (whole paragraph) and insert with trackChanges', async () => {
    const { parser, ast, engine } = await open(PARA + '<w:p><w:r><w:t>second</w:t></w:r></w:p>');
    const [first, second] = ast.sections[0].nodes.map(n => n.id);
    const del = await engine.executeBatch({ operations: [{ action: 'delete', target: first, trackChanges: true }] } as never);
    expect(del.success, JSON.stringify(del.errors)).toBe(true);
    const ins = await engine.executeBatch({ operations: [{ action: 'insert', type: 'paragraph', anchor: second, position: 'after', text: 'added', trackChanges: true }] } as never);
    expect(ins.success, JSON.stringify(ins.errors)).toBe(true);
    const created = [...(del.revisionIds ?? []), ...(ins.revisionIds ?? [])].sort((a, b) => a - b);
    expect(new Set(created).size).toBe(created.length);
    expect(created).toEqual(idsInDocument(parser));
  });

  it('trackedChange insert and delete return the ids they wrote', async () => {
    const { parser, ast, engine } = await open(PARA);
    const first = ast.sections[0].nodes[0].id;
    const ins = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'insert', target: first, text: ' more' }] } as never);
    expect(ins.success, JSON.stringify(ins.errors)).toBe(true);
    const del = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'delete', target: first, text: 'brave' }] } as never);
    expect(del.success, JSON.stringify(del.errors)).toBe(true);
    const created = [...(ins.revisionIds ?? []), ...(del.revisionIds ?? [])].sort((x, y) => x - y);
    expect(created.length).toBeGreaterThanOrEqual(2);
    expect(new Set(created).size).toBe(created.length);
    expect(created).toEqual(idsInDocument(parser));
  });

  it('updateCell on a cell with text', async () => {
    const { parser, ast, engine } = await open(TABLE);
    const table = ast.sections[0].nodes[0].id;
    const r = await engine.executeBatch({ operations: [{ action: 'table', target: table, operation: 'updateCell', rowIndex: 1, colIndex: 1, text: 'ALPHA', trackChanges: true }] } as never);
    expect(r.success, JSON.stringify(r.errors)).toBe(true);
    const created = [...(r.revisionIds ?? [])].sort((a, b) => a - b);
    expect(created.length).toBeGreaterThan(0);
    expect(created).toEqual(idsInDocument(parser));
  });

  it('updateCell on an empty cell', async () => {
    const { parser, ast, engine } = await open(TABLE);
    const table = ast.sections[0].nodes[0].id;
    const r = await engine.executeBatch({ operations: [{ action: 'table', target: table, operation: 'updateCell', rowIndex: 2, colIndex: 2, text: 'filled', trackChanges: true }] } as never);
    expect(r.success, JSON.stringify(r.errors)).toBe(true);
    const created = [...(r.revisionIds ?? [])].sort((a, b) => a - b);
    expect(created.length).toBeGreaterThan(0);
    expect(created).toEqual(idsInDocument(parser));
  });

  it('two tracked edits in one batch never share an id, and the output validates', async () => {
    const { parser, ast, engine } = await open(TABLE + PARA);
    const [table, para] = ast.sections[0].nodes.map(n => n.id);
    const r = await engine.executeBatch({ operations: [
      { action: 'table', target: table, operation: 'updateCell', rowIndex: 1, colIndex: 2, text: 'BETA', trackChanges: true },
      { action: 'replace', target: para, text: 'x y z', preserveFormatting: false, trackChanges: true },
    ] } as never);
    expect(r.success, JSON.stringify(r.errors)).toBe(true);
    const created = r.revisionIds ?? [];
    expect(new Set(created).size).toBe(created.length);
    expect([...created].sort((a, b) => a - b)).toEqual(idsInDocument(parser));
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });
});
