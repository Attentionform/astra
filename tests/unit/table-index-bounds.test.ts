/**
 * A row or column number that is not in the table is refused, for every table
 * operation that takes one. A number from the old 0-based convention (0) used
 * to fall through silently in a few of them: the call said ok and either did
 * nothing or put the change at the end.
 */
import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, type DocumentAST, type Node } from '@attentionform/astra-engine';

interface Fixture {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}

/** Blank doc + a 3-row x 3-col table: header A B C, then 1 2 3, then 4 5 6. */
async function withTable(): Promise<Fixture> {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const inserted = await engine.executeBatch({
    operations: [
      {
        action: 'insert',
        type: 'table',
        anchor: ast.sections[0].nodes[0].id,
        position: 'after',
        headers: ['A', 'B', 'C'],
        rows: [
          ['1', '2', '3'],
          ['4', '5', '6'],
        ],
      },
    ] as never,
  });
  expect(inserted.success).toBe(true);
  return { parser, ast, engine };
}

const tableOf = (f: Fixture): Node => f.ast.sections[0].nodes.find((n) => n.type === 'table') as Node;

const table = (f: Fixture, params: Record<string, unknown>) =>
  f.engine.executeBatch({ operations: [{ action: 'table', target: tableOf(f).id, ...params }] as never });

const shape = (f: Fixture): string[][] => {
  const tbl = tableOf(f).xml as Element;
  return Array.from(tbl.getElementsByTagName('w:tr')).map((tr) =>
    Array.from(tr.childNodes as unknown as Element[])
      .filter((c) => c.tagName === 'w:tc')
      .map((tc) => tc.textContent ?? ''),
  );
};

const gridWidths = (f: Fixture): string[] =>
  Array.from((tableOf(f).xml as Element).getElementsByTagName('w:gridCol')).map((c) => c.getAttribute('w:w') ?? '');

const hasHeaderRepeat = (f: Fixture): boolean => (tableOf(f).xml as Element).getElementsByTagName('w:tblHeader').length > 0;

const BAD_NUMBERS = [0, 4, 99, -1];

describe('addRow afterIndex', () => {
  it.each(BAD_NUMBERS)('afterIndex %i is refused with the valid range and nothing changes', async (afterIndex) => {
    const f = await withTable();
    const before = shape(f);
    const result = await table(f, { operation: 'addRow', afterIndex, cells: ['x', 'y', 'z'] });
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(
      new RegExp(`addRow row index ${afterIndex} out of bounds \\(table has 3 rows\\)\\. Remedy: pass afterIndex in \\[1, 3\\], counting from 1`),
    );
    expect(shape(f)).toEqual(before);
  });

  it('afterIndex 3 (the last row) and an omitted afterIndex still append', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'addRow', afterIndex: 3, cells: ['p', 'q', 'r'] })).success).toBe(true);
    expect((await table(f, { operation: 'addRow', cells: ['s', 't', 'u'] })).success).toBe(true);
    expect(shape(f).map((r) => r[0])).toEqual(['A', '1', '4', 'p', 's']);
  });
});

describe('addColumn afterIndex', () => {
  it.each(BAD_NUMBERS)('afterIndex %i is refused with the valid range and nothing changes', async (afterIndex) => {
    const f = await withTable();
    const before = shape(f);
    const result = await table(f, { operation: 'addColumn', afterIndex });
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(
      new RegExp(`addColumn column index ${afterIndex} out of bounds \\(row has 3 logical columns\\)\\. Remedy: pass afterIndex in \\[1, 3\\], counting from 1`),
    );
    expect(shape(f)).toEqual(before);
    expect(gridWidths(f)).toHaveLength(3);
  });

  it('afterIndex 3 (the last column) and an omitted afterIndex still append', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'addColumn', afterIndex: 3 })).success).toBe(true);
    expect((await table(f, { operation: 'addColumn' })).success).toBe(true);
    expect(shape(f).every((r) => r.length === 5)).toBe(true);
  });
});

describe('setColumnWidth colIndex', () => {
  it.each(BAD_NUMBERS)('colIndex %i is refused with the valid range and nothing changes', async (colIndex) => {
    const f = await withTable();
    const before = gridWidths(f);
    const result = await table(f, { operation: 'setColumnWidth', colIndex, width: 1234 });
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(
      new RegExp(`setColumnWidth column index ${colIndex} out of bounds \\(table has 3 columns\\)\\. Remedy: pass colIndex in \\[1, 3\\], counting from 1`),
    );
    expect(gridWidths(f)).toEqual(before);
  });

  it('colIndex 2 sets the second column; an omitted colIndex sets the first', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'setColumnWidth', colIndex: 2, width: 1234 })).success).toBe(true);
    expect(gridWidths(f)[1]).toBe('1234');
    expect((await table(f, { operation: 'setColumnWidth', width: 4321 })).success).toBe(true);
    expect(gridWidths(f)[0]).toBe('4321');
  });
});

describe('setHeaderRepeat rowIndex', () => {
  it.each(BAD_NUMBERS)('rowIndex %i is refused with the valid range and nothing changes', async (rowIndex) => {
    const f = await withTable();
    const before = hasHeaderRepeat(f);
    const result = await table(f, { operation: 'setHeaderRepeat', rowIndex });
    expect(result.success).toBe(false);
    expect(JSON.stringify(result.errors)).toMatch(
      new RegExp(`setHeaderRepeat row index ${rowIndex} out of bounds \\(table has 3 rows\\)\\. Remedy: pass rowIndex in \\[1, 3\\], counting from 1`),
    );
    expect(hasHeaderRepeat(f)).toBe(before);
  });

  it('rowIndex 1 marks the first row; an omitted rowIndex marks the first row too', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'setHeaderRepeat', rowIndex: 1 })).success).toBe(true);
    expect(hasHeaderRepeat(f)).toBe(true);
  });
});

describe('every other table operation refuses 0 and a number past the end', () => {
  const CASES: Array<{ name: string; row?: boolean; build: (n: number) => Record<string, unknown> }> = [
    { name: 'updateCell rowIndex', build: (n) => ({ operation: 'updateCell', rowIndex: n, colIndex: 1, text: 'x' }) },
    { name: 'updateCell colIndex', build: (n) => ({ operation: 'updateCell', rowIndex: 1, colIndex: n, text: 'x' }) },
    { name: 'deleteRow', build: (n) => ({ operation: 'deleteRow', index: n }) },
    { name: 'deleteColumn', build: (n) => ({ operation: 'deleteColumn', index: n }) },
    { name: 'setCellShading rowIndex', build: (n) => ({ operation: 'setCellShading', rowIndex: n, colIndex: 1, fill: 'FF0000' }) },
    { name: 'setCellShading colIndex', build: (n) => ({ operation: 'setCellShading', rowIndex: 1, colIndex: n, fill: 'FF0000' }) },
    { name: 'setRowHeight', build: (n) => ({ operation: 'setRowHeight', rowIndex: n, height: 600 }) },
    { name: 'setVerticalAlign rowIndex', build: (n) => ({ operation: 'setVerticalAlign', rowIndex: n, colIndex: 1, align: 'center' }) },
    { name: 'setCellAlignment rowIndex', build: (n) => ({ operation: 'setCellAlignment', rowIndex: n, colIndex: 1, align: 'center' }) },
    { name: 'setCellPadding rowIndex', build: (n) => ({ operation: 'setCellPadding', rowIndex: n, colIndex: 1, top: 40 }) },
    { name: 'setCellTextColor rowIndex', build: (n) => ({ operation: 'setCellTextColor', rowIndex: n, colIndex: 1, color: 'FF0000' }) },
    { name: 'splitCell rowIndex', build: (n) => ({ operation: 'splitCell', rowIndex: n, colIndex: 1 }) },
    { name: 'mergeCells fromRow', build: (n) => ({ operation: 'mergeCells', fromRow: n, fromCol: 1, toRow: 2, toCol: 2 }) },
    { name: 'mergeCells toCol', build: (n) => ({ operation: 'mergeCells', fromRow: 1, fromCol: 1, toRow: 2, toCol: n }) },
  ];

  for (const { name, build } of CASES) {
    for (const n of [0, 99]) {
      it(`${name} ${n} is refused`, async () => {
        const f = await withTable();
        const before = shape(f);
        const result = await table(f, build(n));
        expect(result.success, JSON.stringify(result.errors)).toBe(false);
        expect(shape(f)).toEqual(before);
      });
    }
  }
});
