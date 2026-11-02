import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

/**
 * A table written by `insert` is as wide as the text area. It carried no w:tblW, and Word fits such a table to
 * its text (two short columns gave a narrow table) while LibreOffice drew it at the width of the grid.
 */
async function insertTable(extra: Record<string, unknown> = {}): Promise<Element> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const res = await engine.executeBatch({
    operations: [{ action: 'insert', type: 'table', anchor: 'end', headers: ['Name', 'Value'], rows: [['Alpha', '1']], id: 't1', ...extra }] as never,
    sequential: true,
  });
  expect(res.success).toBe(true);
  return ast.sections.flatMap((s) => s.nodes).find((n) => n.id === 't1')!.xml as Element;
}

const kids = (el: Element): string[] =>
  Array.from(el.childNodes).filter((n) => n.nodeType === 1).map((n) => (n as Element).tagName);

describe('inserted table width', () => {
  it('is 100% of the text area (5000 fiftieths of a percent)', async () => {
    const tbl = await insertTable();
    const tblPr = tbl.getElementsByTagName('w:tblPr')[0];
    const tblW = tblPr.getElementsByTagName('w:tblW')[0];
    expect(tblW.getAttribute('w:w')).toBe('5000');
    expect(tblW.getAttribute('w:type')).toBe('pct');
  });

  it('sits in schema order: after the style, before the borders and the cell margins', async () => {
    const tbl = await insertTable();
    expect(kids(tbl.getElementsByTagName('w:tblPr')[0])).toEqual(['w:tblStyle', 'w:tblW', 'w:tblBorders', 'w:tblCellMar']);
  });

  it('keeps schema order when cell padding adds w:tblCellMar', async () => {
    const tbl = await insertTable({ cellPadding: { top: 60, left: 100, bottom: 60, right: 100 } });
    expect(kids(tbl.getElementsByTagName('w:tblPr')[0])).toEqual(['w:tblStyle', 'w:tblW', 'w:tblBorders', 'w:tblCellMar']);
  });

  it('writes the width once, however many tables are inserted', async () => {
    const tbl = await insertTable();
    expect(tbl.getElementsByTagName('w:tblW').length).toBe(1);
  });
});
