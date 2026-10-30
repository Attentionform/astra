import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { compileMarkdown, composeDocument } from '../src/compose/markdown';
import { createWordElement } from '../src/ooxml/ns';
import { canonicalizeOOXML, sortChildrenBySchema, TBLPR_CHILD_ORDER } from '../src/ooxml/schema-order';
import type { Node } from '../src/ast';
import type { InsertOperation } from '../src/operations/insert';

function directChildren(parent: Element, tagName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === tagName) {
      out.push(child as Element);
    }
  }
  return out;
}

async function freshDoc() {
  const parser = new DocxParser(createBlankDocxBuffer('Table Cell Padding Test Document'));
  const ast = await parser.getAST();
  return { parser, ast };
}

function lastNodeId(ast: { sections: Array<{ nodes: Node[] }> }): string {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1].id;
}

function tableNodeOf(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const found = ast.sections[0].nodes.find((n) => n.type === 'table');
  if (!found) throw new Error('test setup: no table node found');
  return found;
}

describe('table baseline cell padding', () => {
  it('inserts table with numeric cellPadding applying uniform DXA to all sides', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNodeId(ast);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Header 1', 'Header 2'],
      rows: [['Val 1', 'Val 2']],
      cellPadding: 150,
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);
    const tblEl = tblNode.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(tblPr).toBeDefined();

    const tblCellMar = directChildren(tblPr, 'w:tblCellMar')[0];
    expect(tblCellMar).toBeDefined();

    const top = directChildren(tblCellMar, 'w:top')[0];
    const left = directChildren(tblCellMar, 'w:left')[0];
    const bottom = directChildren(tblCellMar, 'w:bottom')[0];
    const right = directChildren(tblCellMar, 'w:right')[0];

    expect(top).toBeDefined();
    expect(top.getAttribute('w:w')).toBe('150');
    expect(top.getAttribute('w:type')).toBe('dxa');

    expect(left).toBeDefined();
    expect(left.getAttribute('w:w')).toBe('150');
    expect(left.getAttribute('w:type')).toBe('dxa');

    expect(bottom).toBeDefined();
    expect(bottom.getAttribute('w:w')).toBe('150');
    expect(bottom.getAttribute('w:type')).toBe('dxa');

    expect(right).toBeDefined();
    expect(right.getAttribute('w:w')).toBe('150');
    expect(right.getAttribute('w:type')).toBe('dxa');
  });

  it('inserts table with params-nested numeric cellPadding', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNodeId(ast);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['A', 'B'],
      rows: [['1', '2']],
      params: { cellPadding: 180 },
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);
    const tblEl = tblNode.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    const tblCellMar = directChildren(tblPr, 'w:tblCellMar')[0];
    expect(tblCellMar).toBeDefined();

    const left = directChildren(tblCellMar, 'w:left')[0];
    expect(left).toBeDefined();
    expect(left.getAttribute('w:w')).toBe('180');
    expect(left.getAttribute('w:type')).toBe('dxa');
  });

  it('inserts table with object cellPadding specifying individual sides', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNodeId(ast);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Col 1', 'Col 2'],
      rows: [['Data 1', 'Data 2']],
      cellPadding: { top: 120, bottom: 120, left: 160, right: 160 },
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);
    const tblEl = tblNode.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    const tblCellMar = directChildren(tblPr, 'w:tblCellMar')[0];
    expect(tblCellMar).toBeDefined();

    const top = directChildren(tblCellMar, 'w:top')[0];
    const left = directChildren(tblCellMar, 'w:left')[0];
    const bottom = directChildren(tblCellMar, 'w:bottom')[0];
    const right = directChildren(tblCellMar, 'w:right')[0];

    expect(top).toBeDefined();
    expect(top.getAttribute('w:w')).toBe('120');
    expect(top.getAttribute('w:type')).toBe('dxa');

    expect(left).toBeDefined();
    expect(left.getAttribute('w:w')).toBe('160');
    expect(left.getAttribute('w:type')).toBe('dxa');

    expect(bottom).toBeDefined();
    expect(bottom.getAttribute('w:w')).toBe('120');
    expect(bottom.getAttribute('w:type')).toBe('dxa');

    expect(right).toBeDefined();
    expect(right.getAttribute('w:w')).toBe('160');
    expect(right.getAttribute('w:type')).toBe('dxa');
  });

  it('preserves byte-identical behavior without tblCellMar when cellPadding is omitted', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNodeId(ast);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Col 1', 'Col 2'],
      rows: [['Data 1', 'Data 2']],
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);
    const tblEl = tblNode.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(tblPr).toBeDefined();

    const tblCellMar = directChildren(tblPr, 'w:tblCellMar');
    expect(tblCellMar.length).toBe(0);

    const childNames = Array.from(tblPr.childNodes)
      .filter((n) => n.nodeType === 1)
      .map((n) => (n as Element).tagName);
    expect(childNames).toEqual(['w:tblStyle', 'w:tblBorders']);
  });

  it('emits cellPadding from markdown compiler and renders tblCellMar in document', async () => {
    const markdown = '| Col 1 | Col 2 |\n| --- | --- |\n| Cell 1 | Cell 2 |';
    const ops = compileMarkdown(markdown);

    const tableInsertOp = ops.find(
      (op) => op.action === 'insert' && (op as { type?: string }).type === 'table'
    ) as InsertOperation | undefined;

    expect(tableInsertOp).toBeDefined();
    expect(tableInsertOp!.cellPadding).toEqual({
      top: 120,
      bottom: 120,
      left: 160,
      right: 160,
    });

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, markdown);
    expect(result.success).toBe(true);

    const ast = await parser.getAST();
    const tblNode = tableNodeOf(ast);
    const tblEl = tblNode.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(tblPr).toBeDefined();

    const tblCellMar = directChildren(tblPr, 'w:tblCellMar')[0];
    expect(tblCellMar).toBeDefined();

    const top = directChildren(tblCellMar, 'w:top')[0];
    const left = directChildren(tblCellMar, 'w:left')[0];
    const bottom = directChildren(tblCellMar, 'w:bottom')[0];
    const right = directChildren(tblCellMar, 'w:right')[0];

    expect(top.getAttribute('w:w')).toBe('120');
    expect(top.getAttribute('w:type')).toBe('dxa');
    expect(left.getAttribute('w:w')).toBe('160');
    expect(left.getAttribute('w:type')).toBe('dxa');
    expect(bottom.getAttribute('w:w')).toBe('120');
    expect(bottom.getAttribute('w:type')).toBe('dxa');
    expect(right.getAttribute('w:w')).toBe('160');
    expect(right.getAttribute('w:type')).toBe('dxa');
  });

  it('performs save and re-open round-trip via parser.rezip() and validates OOXML schema', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNodeId(ast);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Col A', 'Col B'],
      rows: [['Val A', 'Val B']],
      cellPadding: { top: 120, bottom: 120, left: 160, right: 160 },
    } as InsertOperation);

    const buffer = parser.rezip();
    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);

    const reparsed = new DocxParser(buffer);
    const reloadedAst = await reparsed.getAST();
    const reloadedTable = tableNodeOf(reloadedAst);
    expect(reloadedTable).toBeDefined();

    const tblEl = reloadedTable.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(tblPr).toBeDefined();

    const tblCellMar = directChildren(tblPr, 'w:tblCellMar')[0];
    expect(tblCellMar).toBeDefined();

    const top = directChildren(tblCellMar, 'w:top')[0];
    const left = directChildren(tblCellMar, 'w:left')[0];
    const bottom = directChildren(tblCellMar, 'w:bottom')[0];
    const right = directChildren(tblCellMar, 'w:right')[0];

    expect(top.getAttribute('w:w')).toBe('120');
    expect(top.getAttribute('w:type')).toBe('dxa');
    expect(left.getAttribute('w:w')).toBe('160');
    expect(left.getAttribute('w:type')).toBe('dxa');
    expect(bottom.getAttribute('w:w')).toBe('120');
    expect(bottom.getAttribute('w:type')).toBe('dxa');
    expect(right.getAttribute('w:w')).toBe('160');
    expect(right.getAttribute('w:type')).toBe('dxa');

    const validation = await reparsed.validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('verifies strict schema order in tblPr where tblCellMar follows tblBorders/tblLayout and precedes tblLook', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = lastNodeId(ast);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Header 1', 'Header 2'],
      rows: [['Data 1', 'Data 2']],
      cellPadding: { top: 120, bottom: 120, left: 160, right: 160 },
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);
    const tblEl = tblNode.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    const doc = tblEl.ownerDocument;

    // Add tblLayout and tblLook to test ordering
    const tblLayout = createWordElement(doc, 'w:tblLayout');
    tblLayout.setAttribute('w:type', 'fixed');
    const tblLook = createWordElement(doc, 'w:tblLook');
    tblLook.setAttribute('w:val', '04A0');

    // Append in arbitrary order
    tblPr.appendChild(tblLook);
    tblPr.appendChild(tblLayout);

    // Apply schema sorting
    sortChildrenBySchema(tblPr, TBLPR_CHILD_ORDER);

    const childNames = Array.from(tblPr.childNodes)
      .filter((n) => n.nodeType === 1)
      .map((n) => (n as Element).tagName);

    const bordersIdx = childNames.indexOf('w:tblBorders');
    const layoutIdx = childNames.indexOf('w:tblLayout');
    const cellMarIdx = childNames.indexOf('w:tblCellMar');
    const lookIdx = childNames.indexOf('w:tblLook');

    expect(bordersIdx).toBeGreaterThan(-1);
    expect(layoutIdx).toBeGreaterThan(-1);
    expect(cellMarIdx).toBeGreaterThan(-1);
    expect(lookIdx).toBeGreaterThan(-1);

    expect(bordersIdx).toBeLessThan(layoutIdx);
    expect(layoutIdx).toBeLessThan(cellMarIdx);
    expect(cellMarIdx).toBeLessThan(lookIdx);
  });

  it('ensures canonicalizeOOXML sets missing w:type="dxa" on tblCellMar child elements', () => {
    const xml =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body>' +
      '<w:tbl>' +
      '<w:tblPr>' +
      '<w:tblCellMar>' +
      '<w:top w:w="120"/>' +
      '<w:left w:w="160"/>' +
      '<w:bottom w:w="120"/>' +
      '<w:right w:w="160"/>' +
      '</w:tblCellMar>' +
      '</w:tblPr>' +
      '</w:tbl>' +
      '</w:body>' +
      '</w:document>';
    const doc = new DOMParser().parseFromString(xml, 'text/xml');

    canonicalizeOOXML(doc);

    const top = doc.getElementsByTagName('w:top')[0];
    const left = doc.getElementsByTagName('w:left')[0];
    const bottom = doc.getElementsByTagName('w:bottom')[0];
    const right = doc.getElementsByTagName('w:right')[0];

    expect(top.getAttribute('w:type')).toBe('dxa');
    expect(left.getAttribute('w:type')).toBe('dxa');
    expect(bottom.getAttribute('w:type')).toBe('dxa');
    expect(right.getAttribute('w:type')).toBe('dxa');
  });

  describe('nested params.cellPadding validation', () => {
    it('rejects string nested params.cellPadding with typed E_VALIDATION', async () => {
      const { parser, ast } = await freshDoc();
      const engine = new OperationEngine(ast, parser);
      const anchor = lastNodeId(ast);

      await expect(
        engine.execute({
          action: 'insert',
          type: 'table',
          anchor,
          position: 'after',
          headers: ['A', 'B'],
          rows: [['1', '2']],
          params: { cellPadding: 'bad' },
        } as InsertOperation)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
      });
    });

    it('rejects empty object nested params.cellPadding with typed E_VALIDATION', async () => {
      const { parser, ast } = await freshDoc();
      const engine = new OperationEngine(ast, parser);
      const anchor = lastNodeId(ast);

      await expect(
        engine.execute({
          action: 'insert',
          type: 'table',
          anchor,
          position: 'after',
          headers: ['A', 'B'],
          rows: [['1', '2']],
          params: { cellPadding: {} },
        } as InsertOperation)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
      });
    });

    it('rejects bad side value in nested params.cellPadding with typed E_VALIDATION', async () => {
      const { parser, ast } = await freshDoc();
      const engine = new OperationEngine(ast, parser);
      const anchor = lastNodeId(ast);

      const badCases = [
        { cellPadding: { top: 'bad' } },
        { cellPadding: { top: -1 } },
        { cellPadding: { left: 1.5 } },
        { cellPadding: { bottom: NaN } },
        { cellPadding: { unknownSide: 100 } },
      ];

      for (const params of badCases) {
        await expect(
          engine.execute({
            action: 'insert',
            type: 'table',
            anchor,
            position: 'after',
            headers: ['A', 'B'],
            rows: [['1', '2']],
            params,
          } as InsertOperation)
        ).rejects.toMatchObject({
          code: 'E_VALIDATION',
        });
      }
    });

    it('ensures no empty tblCellMar is created or appended on invalid cellPadding', async () => {
      const { parser, ast } = await freshDoc();
      const engine = new OperationEngine(ast, parser);
      const anchor = lastNodeId(ast);

      try {
        await engine.execute({
          action: 'insert',
          type: 'table',
          anchor,
          position: 'after',
          headers: ['A', 'B'],
          rows: [['1', '2']],
          params: { cellPadding: {} },
        } as InsertOperation);
        expect.unreachable('Should have thrown');
      } catch (err: unknown) {
        const error = err as { code?: string; message?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.remedy).toBeDefined();
      }

      // Live AST must remain uncorrupted with no table inserted
      const tableNodes = ast.sections[0].nodes.filter((n) => n.type === 'table');
      expect(tableNodes.length).toBe(0);
    });
  });
});
