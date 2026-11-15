/**
 * 0.5.3: setTableWidth and setRowHeights.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer, validateOperationStrict } from '@attentionform/astra-engine';
import { applyTableWidth, containerWidth, parseTableWidth, scaleGrid } from '../../packages/engine/src/operations/table-layout';

async function session() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const zip = () => new AdmZip(parser.rezip(ast));
  const xml = () => zip().readAsText('word/document.xml');
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  const anchor = ast.sections[0].nodes[0].id as string;
  return { run, xml, zip, anchor, parser, ast };
}
const WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const insertTable = (anchor: string, cols = 4, rows = 3, extra: Record<string, unknown> = {}) => ({
  action: 'insert', type: 'table', id: 'tbl_a', anchor, position: 'after',
  headers: Array.from({ length: cols }, (_, i) => `H${i + 1}`),
  rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => `r${r + 1}c${c + 1}`)),
  ...extra,
});
const op = (operation: string, extra: Record<string, unknown> = {}) => ({ action: 'table', target: 'tbl_a', operation, ...extra });
const tblOf = (xml: string) => xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)![0];
const grid = (tbl: string) => [...tbl.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((m) => Number(m[1]));
const tblW = (tbl: string) => tbl.match(/<w:tblW [^>]*\/>/)![0];
const rowCellWidths = (tbl: string) => [...tbl.matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)].map((tr) => [...tr[0].matchAll(/<w:tcW w:w="(\d+)" w:type="(\w+)"\/>/g)].map((m) => `${m[1]}:${m[2]}`));
const sum = (a: number[]) => a.reduce((x, y) => x + y, 0);
async function fails(s: { run: (o: unknown[]) => Promise<any> }, ops: unknown[], re: RegExp) {
  const res = await s.run(ops);
  expect(res.success).toBe(false);
  expect(JSON.stringify(res)).toMatch(re);
}

describe('setTableWidth: scale (default)', () => {
  it('widthIn sets tblW dxa, scales the grid to add up exactly, rewrites every cell width and fixes the layout', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setTableWidth', { widthIn: 5 })]);
    const t = tblOf(s.xml());
    expect(tblW(t)).toBe('<w:tblW w:w="7200" w:type="dxa"/>');
    expect(sum(grid(t))).toBe(7200);
    expect(t).toContain('<w:tblLayout w:type="fixed"/>');
    for (const row of rowCellWidths(t)) expect(row).toEqual(grid(t).map((w) => `${w}:dxa`));
    const docx = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(docx.errors).toEqual([]);
  });

  it('widthTwips and widthPercent (of the text width) work; percent writes pct in fiftieths', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setTableWidth', { widthTwips: 6000 })]);
    expect(tblW(tblOf(s.xml()))).toBe('<w:tblW w:w="6000" w:type="dxa"/>');
    expect(sum(grid(tblOf(s.xml())))).toBe(6000);
    await s.run([op('setTableWidth', { widthPercent: 80 })]);
    const t = tblOf(s.xml());
    expect(tblW(t)).toBe('<w:tblW w:w="4000" w:type="pct"/>');
    expect(sum(grid(t))).toBe(7488); // 80% of the 9360-twip text width of the blank document
  });

  it('keeps the proportions and rounds on three and seven columns without losing a twip', async () => {
    for (const [cols, width] of [[3, 7001], [7, 8333]] as const) {
      const s = await session();
      await s.run([insertTable(s.anchor, cols), op('setColumnWidths', { widths: Array.from({ length: cols }, (_, i) => 1000 + i * 300) }), op('setTableWidth', { widthTwips: width })]);
      const t = tblOf(s.xml());
      expect(sum(grid(t))).toBe(width);
      const base = Array.from({ length: cols }, (_, i) => 1000 + i * 300);
      const expected = scaleGrid(base, width);
      expect(grid(t)).toEqual(expected);
      for (let i = 0; i < cols - 1; i++) expect(Math.abs(expected[i] - (base[i] * width) / sum(base))).toBeLessThan(1);
    }
  });

  it('scale then setColumnWidths: the explicit widths win and stay exact', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setTableWidth', { widthIn: 5 }), op('setColumnWidths', { widths: [1000, 2000, 2000, 2000] })]);
    expect(grid(tblOf(s.xml()))).toEqual([1000, 2000, 2000, 2000]);
  });

  it('mixed units: pct table width, pct/auto/nil cell widths are rewritten as dxa from the grid', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setColumnWidths', { widths: [2000, 2000, 2000, 2000] })]);
    // Make the first row's cells pct/auto/nil/dxa and the table width pct.
    const dom = s.parser.getDocumentDom();
    const tbl = dom.getElementsByTagName('w:tbl')[0];
    const tcs = Array.from(tbl.getElementsByTagName('w:tc')).slice(0, 4);
    const types: Array<[string, string]> = [['pct', '1250'], ['auto', '0'], ['nil', '0'], ['dxa', '1111']];
    tcs.forEach((tc, i) => {
      const w = tc.getElementsByTagName('w:tcW')[0];
      w.setAttribute('w:type', types[i][0]);
      w.setAttribute('w:w', types[i][1]);
    });
    const res = await s.run([op('setTableWidth', { widthIn: 6 })]);
    expect(res.success).toBe(true);
    const t = tblOf(s.xml());
    expect(sum(grid(t))).toBe(8640);
    expect(rowCellWidths(t)[0]).toEqual(grid(t).map((w) => `${w}:dxa`));
  });

  it('a gridSpan cell takes the sum of its columns; hMerge and vMerge cells keep the width of their own column', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor, 4, 3), op('mergeCells', { fromRow: 3, fromCol: 1, toRow: 3, toCol: 2 }), op('mergeCells', { fromRow: 3, fromCol: 4, toRow: 4, toCol: 4 })]);
    // Row 2: one cell spanning columns 1-2 (gridSpan), as Word writes a merge.
    const tbl = s.parser.getDocumentDom().getElementsByTagName('w:tbl')[0];
    const row2 = tbl.getElementsByTagName('w:tr')[1];
    const [a, b] = Array.from(row2.getElementsByTagName('w:tc'));
    const tcPr = a.getElementsByTagName('w:tcPr')[0] ?? a.insertBefore(a.ownerDocument.createElementNS(WNS, 'w:tcPr'), a.firstChild);
    const span = a.ownerDocument.createElementNS(WNS, 'w:gridSpan');
    span.setAttribute('w:val', '2');
    tcPr.appendChild(span);
    row2.removeChild(b);
    const res = await s.run([op('setTableWidth', { widthTwips: 8000 })]);
    expect(res.success).toBe(true);
    const t = tblOf(s.xml());
    const g = grid(t);
    expect(sum(g)).toBe(8000);
    const rows = rowCellWidths(t);
    expect(rows[1][0]).toBe(`${g[0] + g[1]}:dxa`);
    expect(rows[1]).toHaveLength(3);
    // hMerge cells (row 3, columns 1-2) keep their own column; vMerge cells (rows 3-4, column 4) keep column 4
    expect(rows[2].slice(0, 2)).toEqual([`${g[0]}:dxa`, `${g[1]}:dxa`]);
    expect(rows[2][3]).toBe(`${g[3]}:dxa`);
    expect(rows[3][3]).toBe(`${g[3]}:dxa`);
    const docx = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(docx.errors).toEqual([]);
  });

  it('a row whose spans do not add up to the grid is refused and the table is unchanged', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const dom = s.parser.getDocumentDom();
    const tbl = dom.getElementsByTagName('w:tbl')[0];
    const lastRow = Array.from(tbl.getElementsByTagName('w:tr')).pop()!;
    lastRow.removeChild(Array.from(lastRow.getElementsByTagName('w:tc')).pop()!);
    const before = tblOf(s.xml());
    await fails(s, [op('setTableWidth', { widthIn: 5 })], /spans 3 grid columns but the table grid has 4/);
    expect(tblOf(s.xml())).toBe(before);
  });

  it('a missing grid is rebuilt from a row of plain dxa widths; with auto or nil it is refused', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setColumnWidths', { widths: [2000, 2000, 2000, 2000] })]);
    let dom = s.parser.getDocumentDom();
    let tbl = dom.getElementsByTagName('w:tbl')[0];
    tbl.removeChild(tbl.getElementsByTagName('w:tblGrid')[0]);
    const ok = await s.run([op('setTableWidth', { widthTwips: 8000 })]);
    expect(ok.success).toBe(true);
    expect(sum(grid(tblOf(s.xml())))).toBe(8000);

    const t = await session();
    await t.run([insertTable(t.anchor)]);
    dom = t.parser.getDocumentDom();
    tbl = dom.getElementsByTagName('w:tbl')[0];
    tbl.removeChild(tbl.getElementsByTagName('w:tblGrid')[0]);
    for (const w of Array.from(tbl.getElementsByTagName('w:tcW'))) { w.setAttribute('w:type', 'auto'); w.setAttribute('w:w', '0'); }
    const before = tblOf(t.xml());
    await fails(t, [op('setTableWidth', { widthTwips: 8000 })], /grid is missing or not numeric/);
    expect(tblOf(t.xml())).toBe(before);
  });

  it('a floating table is refused', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const tbl = s.parser.getDocumentDom().getElementsByTagName('w:tbl')[0];
    const tblPr = tbl.getElementsByTagName('w:tblPr')[0];
    const f = tblPr.ownerDocument.createElementNS(WNS, 'w:tblpPr');
    f.setAttribute('w:horzAnchor', 'page');
    tblPr.appendChild(f);
    await fails(s, [op('setTableWidth', { widthIn: 5 })], /floating table/);
  });

  it('warns when the table is wider than the text width but still writes it', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const res: any = await s.run([op('setTableWidth', { widthIn: 8 })]);
    expect(res.success).toBe(true);
    expect(res.warnings.join('\n')).toMatch(/layoutWarning: the table is 8 in wide, the text width is 6.5 in/);
    expect(sum(grid(tblOf(s.xml())))).toBe(11520);
  });
});

describe('setTableWidth: keep', () => {
  it('changes only tblW: layout, grid and cell widths are byte-identical', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const before = tblOf(s.xml());
    await s.run([op('setTableWidth', { widthIn: 4, columns: 'keep' })]);
    const after = tblOf(s.xml());
    expect(tblW(after)).toBe('<w:tblW w:w="5760" w:type="dxa"/>');
    expect(after.replace(/<w:tblW [^>]*\/>/, '')).toBe(before.replace(/<w:tblW [^>]*\/>/, ''));
  });

  it('keep with a percent width does not need the container', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    await s.run([op('setTableWidth', { widthPercent: 50, columns: 'keep' })]);
    expect(tblW(tblOf(s.xml()))).toBe('<w:tblW w:w="2500" w:type="pct"/>');
  });
});

describe('setTableWidth: arguments', () => {
  it('needs exactly one width and validates it', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const before = tblOf(s.xml());
    await fails(s, [op('setTableWidth')], /needs a width/);
    await fails(s, [op('setTableWidth', { widthIn: 5, widthTwips: 7200 })], /more than one width \(widthIn, widthTwips\)/);
    await fails(s, [op('setTableWidth', { widthPercent: 0 })], /outside 1 to 100/);
    await fails(s, [op('setTableWidth', { widthPercent: 101 })], /outside 1 to 100/);
    await fails(s, [op('setTableWidth', { widthIn: 23 })], /outside the supported range/);
    await fails(s, [op('setTableWidth', { widthIn: 0 })], /outside the supported range/);
    await fails(s, [op('setTableWidth', { widthIn: 5, columns: 'stretch' })], /columns/);
    expect(() => validateOperationStrict({ action: 'table', target: 'tbl_a', operation: 'setTableWidth', widthIn: 5, height: 3 })).toThrow(/unknown parameter "height"/);
    expect(() => validateOperationStrict({ action: 'table', target: 'tbl_a', operation: 'setTableWidth', widthIn: 5, columns: 'keep' })).not.toThrow();
    expect(tblOf(s.xml())).toBe(before);
  });

  it('accepts the arguments flat or inside params', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const res = await s.run([{ action: 'table', target: 'tbl_a', params: { op: 'setTableWidth', widthIn: 5, columns: 'scale' } }]);
    expect(res.success).toBe(true);
    expect(sum(grid(tblOf(s.xml())))).toBe(7200);
  });

  it('parse helpers: 22 in and 20 twips are inclusive limits', () => {
    const pick = (o: Record<string, unknown>) => (k: string) => o[k];
    expect(parseTableWidth(pick({ widthIn: 22 }))).toEqual({ spec: { kind: 'dxa', twips: 31680 }, columns: 'scale' });
    expect(() => parseTableWidth(pick({ widthTwips: 31681 }))).toThrow(/outside the supported range/);
    expect(parseTableWidth(pick({ widthTwips: 20 })).spec).toEqual({ kind: 'dxa', twips: 20 });
    expect(() => parseTableWidth(pick({ widthTwips: 19 }))).toThrow();
    expect(() => parseTableWidth(pick({ widthTwips: 5.5 }))).toThrow(/whole number/);
    expect(() => parseTableWidth(pick({ widthIn: '5' }))).toThrow(/finite number/);
  });
});

describe('container width: nested table and sections', () => {
  const parse = (xml: string) => new DOMParser().parseFromString(`<w:document xmlns:w="${WNS}"><w:body>${xml}</w:body></w:document>`, 'text/xml');
  const nested = (tcPr: string, outerPr = '') => parse(
    `<w:tbl><w:tblPr>${outerPr}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p/></w:tc>` +
    `<w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/>${tcPr}</w:tcPr>` +
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>`
  );
  const innerOf = (doc: Document) => doc.getElementsByTagName('w:tbl')[1];

  it('is the parent cell width less the default cell margins', () => {
    expect(containerWidth(innerOf(nested('')), null)).toMatchObject({ twips: 5000 - 216, label: 'the width of its cell' });
  });

  it('uses the cell margins of the cell or the table when given', () => {
    expect(containerWidth(innerOf(nested('<w:tcMar><w:left w:w="200" w:type="dxa"/><w:right w:w="300" w:type="dxa"/></w:tcMar>')), null).twips).toBe(4500);
    expect(containerWidth(innerOf(nested('', '<w:tblCellMar><w:left w:w="50" w:type="dxa"/><w:right w:w="50" w:type="dxa"/></w:tblCellMar>')), null).twips).toBe(4900);
  });

  it('widthPercent of a nested table is relative to its cell', () => {
    const doc = nested('');
    const inner = innerOf(doc);
    const rows = Array.from(inner.getElementsByTagName('w:tr'));
    applyTableWidth(inner, rows, { spec: { kind: 'pct', percent: 50 }, columns: 'scale' }, null);
    const g = Array.from(inner.getElementsByTagName('w:gridCol')).map((c) => Number(c.getAttribute('w:w')));
    expect(g[0] + g[1]).toBe(Math.round((5000 - 216) / 2));
    // the outer table is untouched
    expect(Array.from(doc.getElementsByTagName('w:gridCol')).slice(0, 2).map((c) => c.getAttribute('w:w'))).toEqual(['3000', '5000']);
  });

  it('is refused when the container cannot be determined (parent grid missing, several text columns)', () => {
    const noGrid = parse(`<w:tbl><w:tblPr/><w:tr><w:tc><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>`);
    const inner = noGrid.getElementsByTagName('w:tbl')[1];
    expect(containerWidth(inner, null).twips).toBeNull();
    expect(() => applyTableWidth(inner, Array.from(inner.getElementsByTagName('w:tr')), { spec: { kind: 'pct', percent: 50 }, columns: 'scale' }, null)).toThrow(/widthPercent is refused/);
    // widthIn still works there
    expect(() => applyTableWidth(inner, Array.from(inner.getElementsByTagName('w:tr')), { spec: { kind: 'dxa', twips: 900 }, columns: 'scale' }, null)).not.toThrow();
    const body = parse(`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`).getElementsByTagName('w:tbl')[0];
    expect(containerWidth(body, { columns: 2 }).twips).toBeNull();
    expect(containerWidth(body, { orientation: 'landscape', paperSize: 'Letter' }).twips).toBe(12960);
  });

  it('the text width of a body table comes from the section (A4, margins)', () => {
    const body = parse(`<w:tbl><w:tblPr/></w:tbl>`).getElementsByTagName('w:tbl')[0];
    expect(containerWidth(body, { paperSize: 'A4', marginLeft: 1000, marginRight: 1000 }).twips).toBe(11906 - 2000);
  });
});

describe('setRowHeights', () => {
  const heights = (xml: string) => [...tblOf(xml).matchAll(/<w:trHeight [^>]*\/>/g)].map((m) => m[0]);

  it('sets every row, atLeast by default; heightIn and heightTwips', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setRowHeights', { heightIn: 0.3 })]);
    expect(heights(s.xml())).toEqual(Array(4).fill('<w:trHeight w:val="432" w:hRule="atLeast"/>'));
    await s.run([op('setRowHeights', { heightTwips: 500, rule: 'exact' })]);
    expect(heights(s.xml())).toEqual(Array(4).fill('<w:trHeight w:val="500" w:hRule="exact"/>'));
    const docx = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(docx.errors).toEqual([]);
  });

  it('exact warns that content can be clipped; atLeast does not', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const exact: any = await s.run([op('setRowHeights', { heightIn: 0.3, rule: 'exact' })]);
    expect(exact.warnings.join('\n')).toMatch(/layoutWarning: rule "exact"/);
    const atLeast: any = await s.run([op('setRowHeights', { heightIn: 0.3 })]);
    expect(atLeast.warnings.join('\n')).not.toMatch(/layoutWarning/);
  });

  it('auto clears the heights, keeps other row properties and removes an empty trPr', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor), op('setHeaderRepeat', { row: 1 }), op('setRowHeights', { heightIn: 0.3 })]);
    expect(tblOf(s.xml())).toContain('<w:tblHeader');
    await s.run([op('setRowHeights', { rule: 'auto' })]);
    const t = tblOf(s.xml());
    expect(t).not.toContain('<w:trHeight');
    expect(t).toContain('<w:tblHeader');
    // rows that only carried a height have no empty trPr left
    expect(t).not.toMatch(/<w:trPr\s*\/>|<w:trPr><\/w:trPr>/);
  });

  it('a height with auto, both heights, no height, and out-of-range heights are errors that change nothing', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const before = tblOf(s.xml());
    await fails(s, [op('setRowHeights', { rule: 'auto', heightIn: 0.3 })], /takes no height/);
    await fails(s, [op('setRowHeights', { heightIn: 0.3, heightTwips: 400 })], /both heightIn and heightTwips/);
    await fails(s, [op('setRowHeights')], /needs a height/);
    await fails(s, [op('setRowHeights', { heightIn: 0 })], /outside the supported range/);
    await fails(s, [op('setRowHeights', { heightIn: 23 })], /outside the supported range/);
    await fails(s, [op('setRowHeights', { heightTwips: 400, rule: 'sometimes' })], /rule/);
    expect(() => validateOperationStrict({ action: 'table', target: 'tbl_a', operation: 'setRowHeights', heightTwips: 400, widthIn: 3 })).toThrow(/unknown parameter "widthIn"/);
    expect(tblOf(s.xml())).toBe(before);
  });

  it('only the rows of the target table change (a nested table keeps its own rows)', async () => {
    const s = await session();
    await s.run([insertTable(s.anchor)]);
    const dom = s.parser.getDocumentDom();
    const tbl = dom.getElementsByTagName('w:tbl')[0];
    const cell = tbl.getElementsByTagName('w:tc')[0];
    const nestedXml = `<w:tbl xmlns:w="${WNS}"><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
    cell.insertBefore(new DOMParser().parseFromString(nestedXml, 'text/xml').documentElement, cell.lastChild);
    await s.run([op('setRowHeights', { heightIn: 0.3 })]);
    const t = tblOf(s.xml().replace(/<w:tbl>(?=[\s\S]*<w:tbl )/, '<w:tbl>'));
    const outerRows = [...s.xml().matchAll(/<w:trHeight /g)].length;
    expect(outerRows).toBe(4);
    expect(t).toBeTruthy();
  });
});
