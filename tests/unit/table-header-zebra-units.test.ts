/**
 * Table header bold, zebra rows that survive addRow, and spacing units.
 *
 * - A table header could not be bold: the style guide had no setting and a `bold` request on a cell
 *   was silently dropped. Headers are now bold unless the guide or the call says otherwise.
 * - addRow copied the source row's fill, so a striped table lost its pattern at the new row.
 * - `spacingAfter: 12` meant 0.6 pt (twips) with nothing saying so; "12pt" is now accepted.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  createBlankDocxBuffer,
  createDefaultStyleGuide,
  parseStyleGuide,
  parseTwips,
  validateOperationStrict,
} from '@attentionform/astra-engine';

async function session(guide?: unknown) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser, guide ? { styleGuide: guide as any } : undefined);
  const xml = () => new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  const tableId = () => ast.sections[0].nodes.find((n: any) => n.type === 'table').id as string;
  const anchor = ast.sections[0].nodes[0].id as string;
  return { run, xml, tableId, anchor, ast };
}

const insertTable = (anchor: string, extra: Record<string, unknown> = {}) => ({
  action: 'insert', type: 'table', id: 'tbl_a', anchor, position: 'after', headers: ['Item', 'Qty'], rows: [['a', '1'], ['b', '2'], ['c', '3']], ...extra,
});

function rowRuns(xml: string): string[][] {
  const tbl = xml.match(/<w:tbl>[\s\S]*<\/w:tbl>/)![0];
  return [...tbl.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((tr) => [...tr[0].matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map((r) => r[1]));
}
const rowFills = (xml: string): (string | null)[] =>
  [...xml.match(/<w:tbl>[\s\S]*<\/w:tbl>/)![0].matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((tr) => tr[0].match(/<w:shd [^>]*w:fill="([^"]+)"/)?.[1] ?? null);

describe('table header bold', () => {
  it('is bold by default, and only the header row', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const rows = rowRuns(s.xml());
    expect(rows[0].every((r) => r.includes('<w:b/>'))).toBe(true);
    expect(rows.slice(1).flat().some((r) => r.includes('<w:b/>'))).toBe(false);
  });

  it('can be turned off per call', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor, { headerBold: false })]);
    expect(rowRuns(s.xml())[0].some((r) => r.includes('<w:b/>'))).toBe(false);
  });

  it('follows the style guide, and an explicit value beats it', async () => {
    const guide = parseStyleGuide('# Style Guide: t\n\n## Tables\n- Header Bold: false\n').config;
    expect(guide.tables.headerBold).toBe(false);
    const off = await session(guide);
    await off.run([insertTable(off.anchor)]);
    expect(rowRuns(off.xml())[0].some((r) => r.includes('<w:b/>'))).toBe(false);
    const on = await session(guide);
    await on.run([insertTable(on.anchor, { headerBold: true })]);
    expect(rowRuns(on.xml())[0].every((r) => r.includes('<w:b/>'))).toBe(true);
  });

  it('is on in the built-in guide, and refuses a non-boolean', async () => {
    expect(createDefaultStyleGuide().tables.headerBold).toBe(true);
    const s = await session();
    const result: any = await s.run([insertTable(s.anchor, { params: { headerBold: 'yes' } })]);
    expect(result.success).toBe(false);
    expect(JSON.stringify(result)).toMatch(/headerBold must be a boolean/);
  });
});

describe('addRow keeps zebra striping', () => {
  async function striped() {
    const s = await session();
    await s.run([insertTable(s.anchor), { action: 'table', target: 'tbl_a', operation: 'setShading', fill: 'F0F4FF', headerFill: '1E3A8A' }]);
    return s;
  }

  it('continues the pattern at the end', async () => {
    const s = await striped();
    await s.run([{ action: 'table', target: s.tableId(), operation: 'addRow', afterIndex: 4, cells: ['d', '4'] }]);
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'FFFFFF', 'F0F4FF', 'FFFFFF']);
  });

  it('re-stripes the rows below an insert in the middle', async () => {
    const s = await striped();
    await s.run([{ action: 'table', target: s.tableId(), operation: 'addRow', afterIndex: 2, cells: ['x', '9'] }]);
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'FFFFFF', 'F0F4FF', 'FFFFFF']);
  });

  it('leaves a table that is not striped alone', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), { action: 'table', target: 'tbl_a', operation: 'setShading', fill: 'F0F4FF', headerFill: '1E3A8A', zebra: false }]);
    await s.run([{ action: 'table', target: s.tableId(), operation: 'addRow', afterIndex: 4, cells: ['d', '4'] }]);
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'F0F4FF', 'F0F4FF', 'F0F4FF']);
  });
});

describe('deleteRow keeps zebra striping', () => {
  async function striped() {
    const s = await session();
    await s.run([insertTable(s.anchor), { action: 'table', target: 'tbl_a', operation: 'setShading', fill: 'F0F4FF', headerFill: '1E3A8A' }]);
    return s;
  }

  it('re-stripes the rows below the one removed', async () => {
    const s = await striped();
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'FFFFFF', 'F0F4FF']);
    await s.run([{ action: 'table', target: s.tableId(), operation: 'deleteRow', index: 2 }]);
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'FFFFFF']);
    expect(rowRuns(s.xml()).map((r) => r.length)).toHaveLength(3);
  });

  it('leaves the pattern alone when the last row goes', async () => {
    const s = await striped();
    await s.run([{ action: 'table', target: s.tableId(), operation: 'deleteRow', index: 4 }]);
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'FFFFFF']);
  });

  it('leaves a table that is not striped alone', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), { action: 'table', target: 'tbl_a', operation: 'setShading', fill: 'F0F4FF', headerFill: '1E3A8A', zebra: false }]);
    await s.run([{ action: 'table', target: s.tableId(), operation: 'deleteRow', index: 2 }]);
    expect(rowFills(s.xml())).toEqual(['1E3A8A', 'F0F4FF', 'F0F4FF']);
  });

  it('does not re-stripe when the header row is removed', async () => {
    const s = await striped();
    await s.run([{ action: 'table', target: s.tableId(), operation: 'deleteRow', index: 1 }]);
    expect(rowFills(s.xml())).toEqual(['F0F4FF', 'FFFFFF', 'F0F4FF']);
  });
});

describe('spacing and indent units', () => {
  it('parses twips and points', () => {
    expect(parseTwips(120)).toBe(120);
    expect(parseTwips('12pt')).toBe(240);
    expect(parseTwips('0.5pt')).toBe(10);
    expect(parseTwips(' 6 pt ')).toBe(120);
    expect(parseTwips('12px')).toBeUndefined();
    expect(parseTwips('12')).toBeUndefined();
    expect(parseTwips(undefined)).toBeUndefined();
  });

  it('style writes "12pt" as 240 twips and a bare number unchanged', async () => {
    const s = await session();
    await s.run([
      { action: 'style', target: s.anchor, spacingAfter: '12pt', indentLeft: '0.5pt', spacingBefore: 120 },
    ]);
    const para = s.xml().match(/<w:spacing [^>]*>/g)!.join(' ');
    expect(para).toContain('w:after="240"');
    expect(para).toContain('w:before="120"');
    expect(s.xml()).toContain('w:left="10"');
  });

  it('refuses a unit it does not know, with the accepted forms in the message', () => {
    expect(() => validateOperationStrict({ action: 'style', target: 'p_1', spacingAfter: '12px' })).toThrow(/twips.*"12pt"/);
  });

  it('a unit it does not know inside params is refused too, never written as "undefined"', async () => {
    const s = await session();
    for (const [key, bad] of [['spacingAfter', '12'], ['spacingBefore', '12px'], ['indentLeft', 'wide'], ['indentFirstLine', '1in'], ['indentHanging', '-3pt']] as const) {
      const result = await s.run([{ action: 'style', target: s.anchor, params: { [key]: bad } }]);
      expect(result.success, `${key}=${bad}`).toBe(false);
      expect(JSON.stringify(result.errors), `${key}=${bad}`).toMatch(new RegExp(`${key} must be a number of twips.*12pt`));
    }
    expect(s.xml()).not.toContain('undefined');
  });

  it('params take twips and points the same way flat settings do', async () => {
    const s = await session();
    const result = await s.run([{ action: 'style', target: s.anchor, params: { spacingAfter: '12pt', indentLeft: 360 } }]);
    expect(result.success).toBe(true);
    expect(s.xml()).toContain('w:after="240"');
    expect(s.xml()).toContain('w:left="360"');
  });

  it('a divider insert takes points too', async () => {
    const s = await session();
    await s.run([{ action: 'insert', type: 'divider', anchor: s.anchor, position: 'after', spacingBefore: '6pt', spacingAfter: '3pt' }]);
    const sp = s.xml().match(/<w:spacing [^>]*>/g)!.join(' ');
    expect(sp).toContain('w:before="120"');
    expect(sp).toContain('w:after="60"');
  });
});
