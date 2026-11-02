import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { OperationEngine, validateOperationStrict } from '@attentionform/astra-engine';

/**
 * `addRow` takes the text of the new cells as `cells`, flat on the operation or inside `params`. The strict
 * schema used to reject the flat form, which the engine already read.
 */
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function tableAst(): any {
  const xml =
    `<w:tbl xmlns:w="${W}"><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr><w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc></w:tr>` +
    `<w:tr><w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  const el = new DOMParser().parseFromString(xml, 'text/xml').documentElement;
  return {
    metadata: {},
    sections: [{ id: 'sec_root', heading: null, nodes: [{ id: 't_001', type: 'table', xml: el, metadata: {}, version: 1, checksum: 'c' }], pageSetup: {}, version: 1, checksum: 's' }],
    styles: {}, numbering: { abstractNum: {}, concreteNum: {} }, relationships: {}, headers: {}, footers: {}, mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'cells' }, nextChangeId: 1,
  };
}

async function lastRowAfter(op: Record<string, unknown>): Promise<string> {
  const ast = tableAst();
  const result = await new OperationEngine(ast).executeBatch({ operations: [op as any], sequential: true });
  expect(result.errors).toEqual([]);
  expect(result.success).toBe(true);
  const rows = ast.sections[0].nodes[0].xml.getElementsByTagName('w:tr');
  expect(rows.length).toBe(3);
  return new XMLSerializer().serializeToString(rows[2]);
}

describe('addRow cells', () => {
  it('the strict schema accepts cells flat on the operation', () => {
    const r = validateOperationStrict({ action: 'table', target: 't_001', operation: 'addRow', cells: ['x', 'y'] });
    expect(r.ok).toBe(true);
  });

  it('the strict schema still accepts cells inside params', () => {
    const r = validateOperationStrict({ action: 'table', target: 't_001', operation: 'addRow', params: { cells: ['x', 'y'] } });
    expect(r.ok).toBe(true);
  });

  it('the schema rejects a cells value that is not a list of text, and a list that is too long', () => {
    const bad = (cells: unknown) => () => validateOperationStrict({ action: 'table', target: 't_001', operation: 'addRow', cells });
    expect(bad('x')).toThrow(/cells/);
    expect(bad([1, 2])).toThrow(/cells/);
    expect(bad(Array(101).fill('x'))).toThrow(/cells/);
    expect(bad(['x'.repeat(100_001)])).toThrow(/cells/);
  });

  it('flat cells and params.cells give the same row', async () => {
    const flat = await lastRowAfter({ action: 'table', target: 't_001', operation: 'addRow', cells: ['New 1', 'New 2'] });
    const nested = await lastRowAfter({ action: 'table', target: 't_001', operation: 'addRow', params: { cells: ['New 1', 'New 2'] } });
    expect(flat).toContain('New 1');
    expect(flat).toContain('New 2');
    expect(flat).toBe(nested);
  });
});
