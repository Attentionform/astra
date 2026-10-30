/**
 * Rows and columns are numbered from 1 everywhere: in the inspect grid (r1c1),
 * in cell addresses (t_001:r3c2) and in every table parameter (rowIndex,
 * colIndex, columnIndex, row, col, index, afterIndex, fromRow/toRow/fromCol/toCol).
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, composeDocument, type DocumentAST, type Node } from '@attentionform/astra-engine';

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

const cells = (f: Fixture): string[][] => {
  const tbl = tableOf(f).xml as Element;
  return Array.from(tbl.getElementsByTagName('w:tr')).map((tr) =>
    Array.from(tr.childNodes as unknown as Element[])
      .filter((c) => c.tagName === 'w:tc')
      .map((tc) => tc.textContent ?? ''),
  );
};

const table = (f: Fixture, params: Record<string, unknown>) =>
  f.engine.executeBatch({ operations: [{ action: 'table', target: tableOf(f).id, ...params }] as never });

describe('table numbers count from 1', () => {
  it('updateCell rowIndex 1 / colIndex 1 is the top-left cell, as r1c1 is', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'updateCell', rowIndex: 1, colIndex: 1, text: 'TL' })).success).toBe(true);
    expect(cells(f)[0]).toEqual(['TL', 'B', 'C']);
    expect((await table(f, { operation: 'updateCell', rowIndex: 3, colIndex: 3, text: 'BR' })).success).toBe(true);
    expect(cells(f)[2]).toEqual(['4', '5', 'BR']);
  });

  it('a cell address and rowIndex/colIndex with the same numbers reach the same cell', async () => {
    const f = await withTable();
    await table(f, { operation: 'updateCell', rowIndex: 3, colIndex: 2, text: 'via table op' });
    const viaAddress = await f.engine.executeBatch({
      operations: [{ action: 'replace', target: `${tableOf(f).id}:r2c3`, text: 'via address' }] as never,
    });
    expect(viaAddress.success).toBe(true);
    expect(cells(f)).toEqual([
      ['A', 'B', 'C'],
      ['1', '2', 'via address'],
      ['4', 'via table op', '6'],
    ]);
  });

  it('row, col and columnIndex are aliases counted from 1 too', async () => {
    const f = await withTable();
    await table(f, { operation: 'updateCell', row: 2, col: 2, text: 'alias' });
    await table(f, { operation: 'updateCell', rowIndex: 2, columnIndex: 3, text: 'alias2' });
    await table(f, { operation: 'updateCell', params: { rowIndex: 3, colIndex: 1 }, text: 'nested' });
    expect(cells(f)).toEqual([
      ['A', 'B', 'C'],
      ['1', 'alias', 'alias2'],
      ['nested', '5', '6'],
    ]);
  });

  it('a whole-number string counts from 1 as well', async () => {
    const f = await withTable();
    await table(f, { operation: 'updateCell', rowIndex: '2', colIndex: '1', text: 'str' });
    expect(cells(f)[1][0]).toBe('str');
  });

  it('0 is no longer a row or column: it is refused with the valid range', async () => {
    const f = await withTable();
    const row = await table(f, { operation: 'updateCell', rowIndex: 0, colIndex: 1, text: 'x' });
    expect(row.success).toBe(false);
    expect(JSON.stringify(row.errors)).toMatch(/row index 0 out of bounds \(table has 3 rows\)\. Remedy: pass rowIndex in \[1, 3\], counting from 1/);
    const col = await table(f, { operation: 'updateCell', rowIndex: 1, colIndex: 0, text: 'x' });
    expect(col.success).toBe(false);
    expect(JSON.stringify(col.errors)).toMatch(/column index 0 out of bounds/);
    const past = await table(f, { operation: 'updateCell', rowIndex: 4, colIndex: 1, text: 'x' });
    expect(past.success).toBe(false);
    expect(JSON.stringify(past.errors)).toMatch(/row index 4 out of bounds/);
  });

  it('deleteRow index N removes row N; index 1 is the first row', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'deleteRow', index: 2 })).success).toBe(true);
    expect(cells(f)).toEqual([
      ['A', 'B', 'C'],
      ['4', '5', '6'],
    ]);
    expect((await table(f, { operation: 'deleteRow', rowIndex: 1 })).success).toBe(true);
    expect(cells(f)).toEqual([['4', '5', '6']]);
  });

  it('deleteColumn index N removes column N', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'deleteColumn', index: 2 })).success).toBe(true);
    expect(cells(f)).toEqual([
      ['A', 'C'],
      ['1', '3'],
      ['4', '6'],
    ]);
    const out = await table(f, { operation: 'deleteColumn', index: 3 });
    expect(out.success).toBe(false);
    expect(JSON.stringify(out.errors)).toMatch(/column index 3 out of bounds \(row has 2 logical columns\)\. Remedy: pass index in \[1, 2\], counting from 1/);
  });

  it('addRow afterIndex N inserts after row N; the default is after the last', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'addRow', afterIndex: 1, cells: ['x', 'y', 'z'] })).success).toBe(true);
    expect(cells(f).map((r) => r[0])).toEqual(['A', 'x', '1', '4']);
    expect((await table(f, { operation: 'addRow', cells: ['last', 'l', 'l'] })).success).toBe(true);
    expect(cells(f).map((r) => r[0])).toEqual(['A', 'x', '1', '4', 'last']);
  });

  it('addColumn afterIndex N inserts after column N', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'addColumn', afterIndex: 1 })).success).toBe(true);
    const grid = cells(f);
    expect(grid.every((r) => r.length === 4)).toBe(true);
    expect(grid[0][0]).toBe('A');
    expect(grid[0][2]).toBe('B');
    expect(grid[0][3]).toBe('C');
  });

  it('mergeCells takes a 1-based range', async () => {
    const f = await withTable();
    expect((await table(f, { operation: 'mergeCells', fromRow: 2, fromCol: 1, toRow: 3, toCol: 2 })).success).toBe(true);
    const trs = Array.from((tableOf(f).xml as Element).getElementsByTagName('w:tr'));
    const mergeAt = (row: number, col: number): string | null => {
      const tcs = Array.from(trs[row].childNodes as unknown as Element[]).filter((c) => c.tagName === 'w:tc');
      const tcPr = Array.from(tcs[col].childNodes as unknown as Element[]).find((c) => c.tagName === 'w:tcPr');
      const h = tcPr ? Array.from(tcPr.childNodes as unknown as Element[]).find((c) => c.tagName === 'w:hMerge') : undefined;
      return h ? (h.getAttribute('w:val') ?? '') : null;
    };
    // The header row is untouched; rows 2 and 3, columns 1 and 2 are merged (restart at row 2, column 1).
    expect(mergeAt(0, 0)).toBeNull();
    expect(mergeAt(1, 0)).toBe('restart');
    expect(mergeAt(1, 1)).toBe('continue');
    expect(mergeAt(2, 1)).toBe('continue');
    expect(mergeAt(1, 2)).toBeNull();
  });

  it('setCellShading and setRowHeight address the same row and column numbers', async () => {
    const f = await withTable();
    await table(f, { operation: 'setCellShading', rowIndex: 1, colIndex: 1, fill: 'FF0000' });
    await table(f, { operation: 'setRowHeight', rowIndex: 3, height: 600, rule: 'exact' });
    const xml = new AdmZip(f.parser.rezip()).readAsText('word/document.xml');
    const trs = xml.split(/<w:tr(?=[ >])/).slice(1);
    expect(trs[0]).toContain('FF0000');
    expect(trs[1]).not.toContain('FF0000');
    expect(trs[2]).toContain('w:trHeight');
    expect(trs[0]).not.toContain('w:trHeight');
  });
});

describe('messages and warnings use the same numbers', () => {
  it('a low-contrast warning names the cell by its 1-based position', async () => {
    const f = await withTable();
    const result = await table(f, { operation: 'setCellShading', rowIndex: 2, colIndex: 3, fill: '000000' });
    expect(result.success).toBe(true);
    expect(result.warnings.join(' ')).toMatch(/cell \(2,3\) of table/);
    const whole = await table(f, { operation: 'setShading', fill: '000000', headerFill: '000000' });
    expect(whole.success).toBe(true);
    expect(whole.warnings.join(' ')).toMatch(/cell \(1,1\) of table/);
  });
});

describe('compose shades the header row with the 1-based numbers', () => {
  it('every header cell is filled, the first body row is not', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '| Region | Revenue |\n| --- | --- |\n| North | 120 |\n| South | 98 |\n| East | 143 |\n');
    expect(result.success).toBe(true);
    const xml = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    const rows = xml.split(/<w:tr(?=[ >])/).slice(1);
    const filled = (row: string): number => (row.match(/<w:shd [^>]*w:fill="[0-9A-Fa-f]{6}"/g) ?? []).length;
    expect(filled(rows[0])).toBe(2);
    expect(rows[1]).toContain('F8FAFC'); // zebra on the first body row, not the header navy
    expect(rows[0]).toContain('1E3A8A');
    expect(rows[1]).not.toContain('1E3A8A');
  });
});
