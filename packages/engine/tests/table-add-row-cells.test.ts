import { describe, it, expect } from 'vitest';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { validateDocxBuffer } from '../src/ooxml/validator-gate';
import type { Node } from '../src/ast';
import type { InsertOperation } from '../src/operations/insert';
import type { TableOperation } from '../src/operations/table';

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
  const parser = new DocxParser(createBlankDocxBuffer('Table Add Row Cells Test Document'));
  const ast = await parser.getAST();
  return { parser, ast };
}

function tableNodeOf(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const found = ast.sections[0].nodes.find((n) => n.type === 'table');
  if (!found) throw new Error('test setup: no table node found');
  return found;
}

describe('table addRow cells population', () => {
  it('populates cell text when addRow specifies cells array in params', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Header 1', 'Header 2', 'Header 3'],
      rows: [['Val 1', 'Val 2', 'Val 3']],
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);

    const op: TableOperation = {
      action: 'table',
      target: tblNode.id,
      params: {
        op: 'addRow',
        afterIndex: 2,
        cells: ['Alpha', 'Beta', 'Gamma'],
      },
    };

    const res = await engine.execute(op);
    expect(res).toBeDefined();

    const tblEl = tblNode.xml as Element;
    const rows = directChildren(tblEl, 'w:tr');
    expect(rows.length).toBe(3);

    const addedRow = rows[2];
    const cells = directChildren(addedRow, 'w:tc');
    expect(cells.length).toBe(3);

    const texts = cells.map((cell) => {
      const tNodes = cell.getElementsByTagName('w:t');
      return Array.from(tNodes).map((t) => t.textContent || '').join('');
    });

    expect(texts).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(tblNode.metadata.text).toContain('Alpha');
    expect(tblNode.metadata.text).toContain('Beta');
    expect(tblNode.metadata.text).toContain('Gamma');

    const buf = parser.rezip();
    const val = await validateDocxBuffer(buf);
    expect(val.ok).toBe(true);
    expect(val.errors).toHaveLength(0);
  });

  it('populates cell text when cells array is passed as flat operation property', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['Col A', 'Col B'],
      rows: [['Row 1 A', 'Row 1 B']],
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);

    const op: TableOperation = {
      action: 'table',
      target: tblNode.id,
      operation: 'addRow',
      afterIndex: 1,
      cells: ['First Data', 'Second Data'],
    };

    const res = await engine.execute(op);
    expect(res).toBeDefined();

    const tblEl = tblNode.xml as Element;
    const rows = directChildren(tblEl, 'w:tr');
    expect(rows.length).toBe(3);

    const addedRow = rows[1];
    const cells = directChildren(addedRow, 'w:tc');
    const texts = cells.map((cell) => {
      const tNodes = cell.getElementsByTagName('w:t');
      return Array.from(tNodes).map((t) => t.textContent || '').join('');
    });

    expect(texts).toEqual(['First Data', 'Second Data']);

    const buf = parser.rezip();
    const val = await validateDocxBuffer(buf);
    expect(val.ok).toBe(true);
    expect(val.errors).toHaveLength(0);
  });

  it('clears cell text when cells array is not provided', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers: ['H1', 'H2'],
      rows: [['Val 1', 'Val 2']],
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);

    await engine.execute({
      action: 'table',
      target: tblNode.id,
      params: {
        op: 'addRow',
        afterIndex: 2,
      },
    } as TableOperation);

    const tblEl = tblNode.xml as Element;
    const rows = directChildren(tblEl, 'w:tr');
    const addedRow = rows[2];
    const cells = directChildren(addedRow, 'w:tc');
    const texts = cells.map((cell) => {
      const tNodes = cell.getElementsByTagName('w:t');
      return Array.from(tNodes).map((t) => t.textContent || '').join('');
    });

    expect(texts).toEqual(['', '']);
  });
});
