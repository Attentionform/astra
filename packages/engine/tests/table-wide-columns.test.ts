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
  const parser = new DocxParser(createBlankDocxBuffer('Wide Table Verification Document'));
  const ast = await parser.getAST();
  return { parser, ast };
}

function tableNodeOf(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const found = ast.sections[0].nodes.find((n) => n.type === 'table');
  if (!found) throw new Error('test setup: no table node found');
  return found;
}

describe('wide table column bounds and validation', () => {
  it('preserves exact column widths and fixed layout across a wide multi-column table', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;

    const colCount = 8;
    const headers = Array.from({ length: colCount }, (_, i) => `Col ${i + 1}`);
    const rowData = Array.from({ length: colCount }, (_, i) => `Cell ${i + 1}`);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers,
      rows: [rowData],
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);

    // Wide column widths totaling 20,400 DXA (> 14 inches, exceeding standard page width)
    const wideWidths = [2400, 2600, 2500, 2800, 2300, 2700, 2500, 2600];

    const op: TableOperation = {
      action: 'table',
      target: tblNode.id,
      params: {
        op: 'setColumnWidths',
        widths: wideWidths,
      },
    };

    const res = await engine.execute(op);
    expect(res).toBeDefined();

    const tblEl = tblNode.xml as Element;

    // Check tblGrid gridCol bounds
    const tblGrid = directChildren(tblEl, 'w:tblGrid')[0];
    expect(tblGrid).toBeDefined();
    const gridCols = directChildren(tblGrid, 'w:gridCol');
    expect(gridCols.length).toBe(colCount);

    const actualGridWidths = gridCols.map((c) => Number(c.getAttribute('w:w')));
    expect(actualGridWidths).toEqual(wideWidths);

    // Check cell widths (tcPr / tcW) across all rows
    const rows = directChildren(tblEl, 'w:tr');
    expect(rows.length).toBe(2);

    for (const r of rows) {
      const cells = directChildren(r, 'w:tc');
      expect(cells.length).toBe(colCount);
      cells.forEach((tc, colIdx) => {
        const tcPr = directChildren(tc, 'w:tcPr')[0];
        expect(tcPr).toBeDefined();
        const tcW = tcPr.getElementsByTagName('w:tcW')[0];
        expect(tcW).toBeDefined();
        expect(Number(tcW.getAttribute('w:w'))).toBe(wideWidths[colIdx]);
        expect(tcW.getAttribute('w:type')).toBe('dxa');
      });
    }

    // Check fixed layout in tblPr
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(tblPr).toBeDefined();
    const tblLayout = tblPr.getElementsByTagName('w:tblLayout')[0];
    expect(tblLayout).toBeDefined();
    expect(tblLayout.getAttribute('w:type')).toBe('fixed');

    // Verify package serialization and schema validity
    const buf = parser.rezip();
    const val = await validateDocxBuffer(buf);
    expect(val.ok).toBe(true);
    expect(val.errors).toHaveLength(0);
  });

  it('adding row and updating cells on a wide table preserves schema validity and column widths', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;

    const colCount = 6;
    const headers = Array.from({ length: colCount }, (_, i) => `H${i + 1}`);
    const rowData = Array.from({ length: colCount }, (_, i) => `D${i + 1}`);

    await engine.execute({
      action: 'insert',
      type: 'table',
      anchor,
      position: 'after',
      headers,
      rows: [rowData],
    } as InsertOperation);

    const tblNode = tableNodeOf(ast);
    const wideWidths = [3000, 3200, 2800, 3500, 2900, 3100];

    await engine.execute({
      action: 'table',
      target: tblNode.id,
      params: {
        op: 'setColumnWidths',
        widths: wideWidths,
      },
    } as TableOperation);

    // Add new row with cells
    const newRowCells = Array.from({ length: colCount }, (_, i) => `NewRowVal${i + 1}`);
    await engine.execute({
      action: 'table',
      target: tblNode.id,
      params: {
        op: 'addRow',
        afterIndex: 2,
        cells: newRowCells,
      },
    } as TableOperation);

    // Update a cell in the wide table
    await engine.execute({
      action: 'table',
      target: tblNode.id,
      params: {
        op: 'updateCell',
        rowIndex: 3,
        colIndex: 4,
        text: 'Updated Wide Cell',
      },
    } as TableOperation);

    const tblEl = tblNode.xml as Element;
    const rows = directChildren(tblEl, 'w:tr');
    expect(rows.length).toBe(3);

    // Verify added row cells inherited tcW matching the columns
    const addedRow = rows[2];
    const cells = directChildren(addedRow, 'w:tc');
    cells.forEach((tc, colIdx) => {
      const tcPr = directChildren(tc, 'w:tcPr')[0];
      const tcW = tcPr.getElementsByTagName('w:tcW')[0];
      expect(Number(tcW.getAttribute('w:w'))).toBe(wideWidths[colIdx]);
    });

    const buf = parser.rezip();
    const val = await validateDocxBuffer(buf);
    expect(val.ok).toBe(true);
    expect(val.errors).toHaveLength(0);
  });
});
