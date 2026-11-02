import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

/**
 * A table written by `insert` had no space above and below the text in its rows (the Normal Table style sets 0),
 * so the text touched the lines. It now gets 3 pt (60 twips) above and below. The sides are left to the style, and
 * a caller that passes cellPadding gets exactly what it asked for.
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

const margins = (tbl: Element): Record<string, string | null> => {
  const mar = tbl.getElementsByTagName('w:tblCellMar')[0];
  const out: Record<string, string | null> = {};
  if (!mar) return out;
  for (const side of ['top', 'left', 'bottom', 'right']) {
    const el = mar.getElementsByTagName(`w:${side}`)[0];
    if (el) out[side] = el.getAttribute('w:w');
  }
  return out;
};

describe('default padding of an inserted table', () => {
  it('is 60 twips above and below, and nothing on the sides', async () => {
    expect(margins(await insertTable())).toEqual({ top: '60', bottom: '60' });
  });

  it('is replaced as a whole by a cellPadding the caller passes', async () => {
    expect(margins(await insertTable({ cellPadding: { top: 120, bottom: 120, left: 160, right: 160 } }))).toEqual({ top: '120', left: '160', bottom: '120', right: '160' });
  });

  it('a number applies to every side', async () => {
    const m = margins(await insertTable({ cellPadding: 100 }));
    expect(m).toEqual({ top: '100', left: '100', bottom: '100', right: '100' });
  });

  it('there is one w:tblCellMar', async () => {
    const tbl = await insertTable();
    expect(tbl.getElementsByTagName('w:tblCellMar').length).toBe(1);
  });
});
