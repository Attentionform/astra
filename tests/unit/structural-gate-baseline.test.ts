/**
 * The structural gate (B-1) checks table, drawing and chart content after a batch. A fault that the
 * document already had before the batch does not refuse the batch; a fault the batch adds does.
 */
import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { OperationEngine } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const GOOD_TABLE = `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc></w:tr>` +
  `<w:tr><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
// The first cell has no paragraph, which the gate refuses.
const BROKEN_TABLE = `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr/></w:tc><w:tc><w:p><w:r><w:t>X2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

function node(id: string, xmlStr: string): any {
  const el = new DOMParser().parseFromString(xmlStr, 'text/xml').documentElement;
  return { id, type: 'table', xml: el, metadata: { text: '' }, version: 1, checksum: 'c' };
}

function astWith(nodes: any[]): any {
  return {
    metadata: {},
    sections: [{ id: 'sec_root', heading: null, nodes, pageSetup: {}, version: 1, checksum: 's' }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'gate' },
    nextChangeId: 1,
  };
}

const update = (target: string, row: number, col: number, text: string) =>
  ({ action: 'table', target, params: { op: 'updateCell', rowIndex: row, colIndex: col, text } }) as any;

describe('structural gate with a fault the document already had', () => {
  it('an edit to a good table is allowed when another table was already broken', async () => {
    const ast = astWith([node('t_bad', BROKEN_TABLE), node('t_good', GOOD_TABLE)]);
    const result = await new OperationEngine(ast).executeBatch({ operations: [update('t_good', 1, 1, 'changed')], sequential: true });
    expect(result.errors).toEqual([]);
    expect(result.success).toBe(true);
    expect(new XMLSerializer().serializeToString(ast.sections[0].nodes[1].xml)).toContain('changed');
  });

  it('the broken table itself is left as it was', async () => {
    const ast = astWith([node('t_bad', BROKEN_TABLE), node('t_good', GOOD_TABLE)]);
    const before = new XMLSerializer().serializeToString(ast.sections[0].nodes[0].xml);
    await new OperationEngine(ast).executeBatch({ operations: [update('t_good', 1, 1, 'changed')], sequential: true });
    expect(new XMLSerializer().serializeToString(ast.sections[0].nodes[0].xml)).toBe(before);
  });

  it('an edit that breaks the good table is still refused and rolled back', async () => {
    const ast = astWith([node('t_bad', BROKEN_TABLE), node('t_good', GOOD_TABLE)]);
    // Delete both rows of the good table: it is left with no rows.
    const result = await new OperationEngine(ast).executeBatch({
      operations: [
        { action: 'table', target: 't_good', params: { op: 'deleteRow', index: 1 } },
        { action: 'table', target: 't_good', params: { op: 'deleteRow', index: 1 } },
      ] as any,
      sequential: true,
    });
    expect(result.success).toBe(false);
    expect(result.rollbackTriggered).toBe(true);
    expect(result.errors.some((e) => String(e.error).includes("Table node 't_good' has no w:tr rows"))).toBe(true);
    expect(ast.sections[0].nodes[1].xml.getElementsByTagName('w:tr').length).toBe(2);
  });

  it('a second fault of another kind in the already broken table is refused', async () => {
    const ast = astWith([node('t_bad', BROKEN_TABLE), node('t_good', GOOD_TABLE)]);
    // The broken cell had no paragraph. An edit that also sets a gridSpan of 0 on the good table is a new fault.
    const goodCell = ast.sections[0].nodes[1].xml.getElementsByTagName('w:tc')[0] as Element;
    const op = update('t_good', 1, 1, 'x');
    const engine = new OperationEngine(ast);
    const spy = engine as unknown as { routeOperation: (o: unknown, d?: string) => Promise<unknown> };
    const original = spy.routeOperation.bind(engine);
    spy.routeOperation = async (o: unknown, d?: string) => {
      const out = await original(o, d);
      const tcPr = goodCell.ownerDocument!.createElementNS(W, 'w:tcPr');
      const gs = goodCell.ownerDocument!.createElementNS(W, 'w:gridSpan');
      gs.setAttribute('w:val', '0');
      tcPr.appendChild(gs);
      goodCell.insertBefore(tcPr, goodCell.firstChild);
      return out;
    };
    const result = await engine.executeBatch({ operations: [op], sequential: true });
    expect(result.success).toBe(false);
    expect(result.errors.some((e) => String(e.error).includes("gridSpan w:val='0'"))).toBe(true);
  });

  it('a batch that makes a new broken table is refused', async () => {
    const ast = astWith([node('t_good', GOOD_TABLE)]);
    const goodCell = ast.sections[0].nodes[0].xml.getElementsByTagName('w:tc')[1] as Element;
    const engine = new OperationEngine(ast);
    const spy = engine as unknown as { routeOperation: (o: unknown, d?: string) => Promise<unknown> };
    const original = spy.routeOperation.bind(engine);
    spy.routeOperation = async (o: unknown, d?: string) => {
      const out = await original(o, d);
      // The edit leaves a cell without a paragraph.
      for (const p of Array.from(goodCell.getElementsByTagName('w:p'))) goodCell.removeChild(p);
      return out;
    };
    const result = await engine.executeBatch({ operations: [update('t_good', 1, 1, 'x')], sequential: true });
    expect(result.success).toBe(false);
    expect(result.errors.some((e) => String(e.error).includes('w:tc with no w:p paragraph'))).toBe(true);
  });

  it('a document with no fault behaves as before', async () => {
    const ast = astWith([node('t_good', GOOD_TABLE)]);
    const result = await new OperationEngine(ast).executeBatch({ operations: [update('t_good', 2, 2, 'ok')], sequential: true });
    expect(result.success).toBe(true);
  });
});
