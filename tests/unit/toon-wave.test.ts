import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  Chart2DManipulator,
  Chart3DManipulator,
  DocxParser,
  ToonGenerator,
  TOON_FLAG_LEGEND,
  collectToonTokens,
  extractTableGrid,
  renderChartText,
  renderTableGrid,
  summarize2DChartXml,
  summarize3DChartXml,
  toonFlagLegendKey,
  type DocumentAST,
  type Node as AstNode,
} from '@attentionform/astra-engine';
import { LEAN6_DESCRIPTIONS } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function para(inner: string): string {
  return `<w:p>${inner}</w:p>`;
}

function run(text: string, rPr = ''): string {
  return `<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

function cell(text: string, tcPr = ''): string {
  return `<w:tc><w:tcPr>${tcPr}<w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${para(run(text))}</w:tc>`;
}

/** Realistic merged table: hMerge restart + explicit continue (Word-style). */
function hMergeTable(): string {
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:noHBand="0"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>` +
    `<w:tr>${cell('Name')}${cell('Q1')}${cell('Q2')}</w:tr>` +
    `<w:tr>${cell('Alpha', '<w:hMerge w:val="restart"/>')}${cell('', '<w:hMerge w:val="continue"/>')}${cell('9')}</w:tr>` +
    `<w:tr>${cell('Beta')}${cell('3')}${cell('4')}</w:tr></w:tbl>`
  );
}

/** gridSpan table: one wide cell covering two logical columns. */
function spanTable(): string {
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>` +
    `<w:tr>${cell('H1')}${cell('H2')}${cell('H3')}</w:tr>` +
    `<w:tr>${cell('Wide', '<w:gridSpan w:val="2"/>')}${cell('9')}</w:tr></w:tbl>`
  );
}

function kitchenSinkBody(): string {
  const styled = para(
    `<w:pPr><w:jc w:val="center"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>` +
      run('Styled', '<w:rPr><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="28"/><w:rStyle w:val="Emphasis"/></w:rPr>')
  );
  const link = para(
    `<w:hyperlink r:id="rId9"><w:r><w:t>linked text</w:t></w:r></w:hyperlink>`
  );
  const bookmarked = `<w:p><w:bookmarkStart w:id="0" w:name="bm0"/><w:r><w:t>marked</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>`;
  const commented = para(`<w:r><w:t>noted</w:t></w:r><w:r><w:commentReference w:id="0"/></w:r>`);
  const cited = para(
    `<w:r><w:fldSimple w:instr=" CITATION Smith00 "><w:r><w:t>CITE Smith 2020</w:t></w:r></w:fldSimple></w:r>`
  );
  const tracked = `<w:p><w:ins w:author="t" w:date="2026-01-01T00:00:00Z"><w:r><w:t>added text</w:t></w:r></w:ins></w:p>`;
  const equation = `<w:p><m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>`;
  const image = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></wp:inline></w:drawing></w:r></w:p>`;
  return (
    para(`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('Kitchen Sink')}`) +
    styled + link + bookmarked + commented + cited + tracked + equation + image +
    hMergeTable() + spanTable() +
    `<w:sectPr/>`
  );
}

function buildKitchenSinkDocx(): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
      'utf8'
    )
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      'utf8'
    )
  );
  const docXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:m="${M_NS}"><w:body>${kitchenSinkBody()}</w:body></w:document>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile(
    'word/styles.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
      'utf8'
    )
  );
  return zip.toBuffer();
}

function tableRowText(ast: DocumentAST, tableId: string, fidelity: 'compact' | 'standard' = 'standard'): { text: string; style: string; flags: string } {
  const gen = new ToonGenerator();
  const table = gen.toon(ast, { fidelity });
  const row = table.rows.find((r) => r.id === tableId);
  if (!row) throw new Error(`table row ${tableId} missing`);
  return { text: row.text, style: row.style, flags: row.flags };
}

function cellAt(toonText: string, coord: string): string | null {
  const m = toonText.match(new RegExp(`${coord} ([^|;\\n]+)`));
  return m ? (m[1] as string).trim() : null;
}

describe('Wave-2D grid TOON (P1-2)', () => {
  it('table rows render delimited cells with exact coordinates', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const { text, flags } = tableRowText(ast, 't_001');
    expect(text).toBe(
      'r1c1 Name | r1c2 Q1 | r1c3 Q2 ; r2c1 Alpha | r2c2 hidden-cover | r2c3 9 ; r3c1 Beta | r3c2 3 | r3c3 4'
    );
    expect(flags).toContain('[TBL]');
  });

  it('agents address exact cells from TOON alone (no XML re-read)', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const { text } = tableRowText(ast, 't_001');
    // Simulated agent: parse coordinates straight out of the TOON text.
    expect(cellAt(text, 'r1c1')).toBe('Name');
    expect(cellAt(text, 'r2c3')).toBe('9');
    expect(cellAt(text, 'r3c2')).toBe('3');
    expect(cellAt(text, 'r2c2')).toBe('hidden-cover');
  });

  it('hMerge covers are tagged hidden-cover (never leak cover text as values)', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const { text } = tableRowText(ast, 't_001');
    expect(text).toContain('r2c2 hidden-cover');
    expect(text).not.toMatch(/r2c2 [^h]/);
  });

  it('gridSpan origins are marked (spanN) and later cells skip covered columns', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const { text } = tableRowText(ast, 't_002');
    expect(text).toBe('r1c1 H1 | r1c2 H2 | r1c3 H3 ; r2c1 Wide(span2) | r2c3 9');
  });

  it('compact stays terse: coordinates present, no standard-only suffix', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const compact = tableRowText(ast, 't_001', 'compact');
    expect(compact.text).toContain('r1c1 Name');
    expect(compact.text).toContain('hidden-cover');
    expect(compact.flags).toBe('[TBL]');
  });

  it('table style token carries live grid dims (table:RxC)', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    expect(tableRowText(ast, 't_001').style).toBe('table:3x3');
    expect(tableRowText(ast, 't_002').style).toBe('table:2x3');
  });

  it('extractTableGrid tags bare hMerge and vMerge covers (ECMA bare-means-continue)', () => {
    const doc = new DOMParser().parseFromString(
      `<w:tbl xmlns:w="${W_NS}"><w:tr>` +
        `<w:tc><w:tcPr><w:hMerge/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>z</w:t></w:r></w:p></w:tc>` +
        `</w:tr></w:tbl>`,
      'text/xml'
    );
    const grid = extractTableGrid(doc.documentElement as unknown as Element);
    expect(grid.rows).toBe(1);
    expect(grid.cols).toBe(3);
    expect(grid.cells[0]?.cover).toBe(true);
    expect(grid.cells[1]?.cover).toBe(false);
    expect(grid.cells[2]?.cover).toBe(true);
  });

  it('nested w:tbl never pollutes outer cell text/span (direct-kids scoping)', () => {
    const doc = new DOMParser().parseFromString(
      `<w:tbl xmlns:w="${W_NS}"><w:tr>` +
        `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>Outer</w:t></w:r></w:p>` +
        `<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        `</w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>` +
        `</w:tr></w:tbl>`,
      'text/xml'
    );
    const grid = extractTableGrid(doc.documentElement as unknown as Element);
    expect(grid.rows).toBe(1);
    expect(grid.cols).toBe(2);
    // Outer cell shows only its own text; nested span must not leak.
    expect(grid.cells[0]?.text).toBe('Outer');
    expect(grid.cells[0]?.span).toBe(1);
    expect(grid.cells[1]?.text).toBe('B');
    expect(renderTableGrid(grid)).toBe('r1c1 Outer | r1c2 B');
  });
});

describe('Wave-2D chart visibility (P1-3)', () => {
  const DEMO = {
    title: 'Revenue by Quarter',
    categories: ['Q1', 'Q2'],
    series: [{ name: 'ARR', values: [10, 25], color: '10B981' }],
  };

  async function chartAst(chartType: 'column' | 'bar3D', title: string) {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    if (chartType === 'bar3D') {
      new Chart3DManipulator(ast, parser).insert3DChart({ chartType, title, categories: DEMO.categories, series: DEMO.series });
    } else {
      new Chart2DManipulator(ast, parser).insert2DChart({ chartType, title, categories: DEMO.categories, series: DEMO.series });
    }
    return new DocxParser(parser.rezip()).getAST();
  }

  it('2D chart TOON row surfaces title + categories + series (read-only)', async () => {
    const ast = await chartAst('column', DEMO.title);
    const gen = new ToonGenerator();
    const table = gen.toon(ast, { fidelity: 'standard' });
    const row = table.rows.find((r) => r.style === 'chart');
    expect(row).toBeDefined();
    expect(row?.text).toContain('[Chart column] Revenue by Quarter');
    expect(row?.text).toContain('cats: Q1,Q2');
    expect(row?.text).toContain('ARR: 10,25');
    expect(row?.flags).toMatch(/\[CH:column\]/);
  });

  it('3D chart TOON row surfaces title + categories + series with chart3d style', async () => {
    const ast = await chartAst('bar3D', 'Pipeline 3D');
    const gen = new ToonGenerator();
    const table = gen.toon(ast, { fidelity: 'standard' });
    const row = table.rows.find((r) => r.style === 'chart3d');
    expect(row).toBeDefined();
    expect(row?.text).toContain('Pipeline 3D');
    expect(row?.text).toContain('cats: Q1,Q2');
    expect(row?.text).toContain('ARR: 10,25');
  });

  it('chart part without title/cats renders explicit-unknown, never placeholders', async () => {
    // Real insert roundtrip, then strip title + series from the stored chart
    // part so the reload hydrates a title-less, cat-less, series-less part.
    const parser = DocxParser.createNewDocument('blank');
    const ast0 = await parser.getAST();
    new Chart2DManipulator(ast0, parser).insert2DChart({
      chartType: 'column', title: DEMO.title, categories: DEMO.categories, series: DEMO.series,
    });
    const zip = new AdmZip(parser.rezip());
    const part = 'word/charts/chart1.xml';
    const xml = zip.readAsText(part);
    expect(xml).toContain('<c:title>');
    const stripped = xml
      .replace(/<c:title>[\s\S]*?<\/c:title>/, '')
      .replace(/<c:ser>[\s\S]*?<\/c:ser>/g, '');
    expect(stripped).not.toContain('<c:title>');
    expect(stripped).not.toContain('<c:ser>');
    zip.updateFile(part, Buffer.from(stripped, 'utf8'));
    const ast = await new DocxParser(zip.toBuffer()).getAST();
    const node = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart');
    expect(node).toBeDefined();
    // Provenance: empty stays empty with synthetic:true — no fabricated text.
    const data = (node?.metadata as { chartData?: Record<string, unknown> }).chartData;
    expect(data?.synthetic).toBe(true);
    expect(data?.title).toBeUndefined();
    expect(data?.categories).toEqual([]);
    expect(data?.series).toEqual([]);
    const gen = new ToonGenerator();
    const table = gen.toon(ast, { fidelity: 'standard' });
    const row = table.rows.find((r) => r.style === 'chart');
    expect(row).toBeDefined();
    expect(row?.text).toContain('[Chart ?]');
    expect(row?.flags).toMatch(/\[CH:\?\]/);
    expect(row?.text).not.toContain('Chart Analysis');
    expect(row?.text).not.toContain('Requirements');
    expect(row?.text).not.toContain('Hours Saved');
    expect(renderChartText(node as AstNode)).toBe('[Chart ?]');
  });

  it('summarize2DChartXml parses real generated chart XML exactly', () => {
    const ast = { sections: [] } as unknown as DocumentAST;
    const xml = new Chart2DManipulator(ast).generate2DChartXml({
      chartType: 'column',
      title: DEMO.title,
      categories: DEMO.categories,
      series: DEMO.series,
    });
    const summary = summarize2DChartXml(xml);
    expect(summary.title).toBe(DEMO.title);
    expect(summary.categories).toEqual(DEMO.categories);
    expect(summary.series).toEqual([{ name: 'ARR', values: [10, 25] }]);
    expect(summary.chartType).toBe('column');
  });

  it('summarize3DChartXml parses real generated 3D XML exactly', () => {
    const ast = { sections: [] } as unknown as DocumentAST;
    const xml = new Chart3DManipulator(ast).generate3DChartXml({
      chartType: 'bar3D',
      title: 'Pipeline 3D',
      categories: DEMO.categories,
      series: DEMO.series,
    });
    const summary = summarize3DChartXml(xml);
    expect(summary.title).toBe('Pipeline 3D');
    expect(summary.categories).toEqual(DEMO.categories);
    expect(summary.series).toEqual([{ name: 'ARR', values: [10, 25] }]);
    expect(summary.chartType).toBe('bar3D');
  });

  it('summarizers never throw on malformed XML (empty summary)', () => {
    expect(summarize2DChartXml('not <xml at all')).toEqual({ title: undefined, categories: [], series: [], chartType: undefined });
    expect(summarize3DChartXml('')).toEqual({ title: undefined, categories: [], series: [], chartType: undefined });
  });
});

describe('Wave-2D docs honesty (P1-7)', () => {
  it('every emitted flag/marker has a legend entry (programmatic completeness)', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const gen = new ToonGenerator();
    // Cover the list-level branch synthetically.
    const doc = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W_NS}"><w:r><w:t>listed</w:t></w:r></w:p>`,
      'text/xml'
    );
    const listNode: AstNode = {
      id: 'syn_list',
      type: 'paragraph',
      xml: doc.documentElement as unknown as Element,
      metadata: {
        text: 'listed', hasCitation: false, hasComment: false, hasBookmark: false,
        hasTrackedChanges: false, hasEquation: false, hasField: false,
        hasImage: false, hasHyperlink: false, listLevel: 2,
      },
      version: 1,
      checksum: 'syn',
    };
    ast.sections[0].nodes.push(listNode);
    // REAL chart node via insert roundtrip so CH:/RID: tokens emit
    // non-vacuously (never hand-forced chartType).
    const cParser = DocxParser.createNewDocument('blank');
    const cAst0 = await cParser.getAST();
    new Chart2DManipulator(cAst0, cParser).insert2DChart({
      chartType: 'column', title: 'Legend Rev', categories: ['Q1', 'Q2'], series: [{ name: 'ARR', values: [10, 25] }],
    });
    const cAst = await new DocxParser(cParser.rezip()).getAST();
    const realChart = cAst.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart');
    expect(realChart).toBeDefined();
    if (realChart) ast.sections[0].nodes.push(realChart);
    // REAL explicit-unknown chart: strip title/series from a second insert so
    // the parser hydrates synthetic:true end-to-end (no hand-forced type).
    const uParser = DocxParser.createNewDocument('blank');
    const uAst0 = await uParser.getAST();
    new Chart2DManipulator(uAst0, uParser).insert2DChart({
      chartType: 'column', title: 'Strip Me', categories: ['Q1'], series: [{ name: 'S', values: [1] }],
    });
    const uZip = new AdmZip(uParser.rezip());
    const uXml = uZip.readAsText('word/charts/chart1.xml')
      .replace(/<c:title>[\s\S]*?<\/c:title>/, '')
      .replace(/<c:ser>[\s\S]*?<\/c:ser>/g, '');
    uZip.updateFile('word/charts/chart1.xml', Buffer.from(uXml, 'utf8'));
    const uAst = await new DocxParser(uZip.toBuffer()).getAST();
    const unknownChart = uAst.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart');
    expect(unknownChart).toBeDefined();
    if (unknownChart) ast.sections[0].nodes.push(unknownChart);
    const standard = gen.toon(ast, { fidelity: 'standard', useCache: false });
    const rendered = gen.renderToonText(standard);
    // Diff markers ride the same legend.
    const oldTable = gen.toon(ast, { fidelity: 'standard', useCache: false });
    const addedRow = { ...standard.rows[0], id: 'syn_added', text: 'fresh' };
    const changedRow = { ...standard.rows[1], text: 'edited body' };
    const newTable = { ...standard, rows: [...standard.rows.slice(0, 1), changedRow, ...standard.rows.slice(2), addedRow] };
    const diff = gen.generateDiff(oldTable, newTable);
    const diffRendered = gen.renderToonText(diff);
    const tokens = [...collectToonTokens(rendered), ...collectToonTokens(diffRendered)];
    expect(tokens.length).toBeGreaterThan(15);
    // Non-vacuous: real chart tokens actually emitted (not just legend keys).
    expect(rendered).toMatch(/\[CH:column\]/);
    expect(rendered).toMatch(/\[RID:/);
    expect(rendered).toMatch(/\[CH:\?\]/);
    expect(rendered).toContain('[Chart ?]');
    const missing = tokens
      .map((t) => ({ token: t, key: toonFlagLegendKey(t) }))
      .filter(({ key }) => !(key in TOON_FLAG_LEGEND));
    expect(missing).toEqual([]);
    // Spot-check the codes the task calls out by name.
    for (const code of ['TBL', 'C:', 'SZ:', 'RS:', 'CH:', 'RID:', 'span', 'hidden-cover', '+ADDED', '~MODIFIED']) {
      expect(Object.keys(TOON_FLAG_LEGEND)).toContain(code);
    }
  });

  it('realistic ids: paragraphs are 8-hex, tables t_NNN, section sec_root', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    expect(ast.sections.map((s) => s.id)).toContain('sec_root');
    const paras = ast.sections[0].nodes.filter((n) => n.type === 'paragraph');
    expect(paras.length).toBeGreaterThan(0);
    for (const p of paras) expect(p.id).toMatch(/^[0-9A-F]{8}$/);
    expect(ast.sections[0].nodes.filter((n) => n.type === 'table').map((n) => n.id)).toEqual(['t_001', 't_002']);
  });

  it('full fidelity without sectionId throws with remedy', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const gen = new ToonGenerator();
    expect(() => gen.toon(ast, { fidelity: 'full' })).toThrow(/sectionId.*Remedy/);
  });

  it('full fidelity with unknown sectionId names the available section ids', async () => {
    const ast = await new DocxParser(buildKitchenSinkDocx()).getAST();
    const gen = new ToonGenerator();
    expect(() => gen.toon(ast, { fidelity: 'full', sectionId: 'sec_nope' })).toThrow(/sec_root/);
  });

  it('inspect description documents grid, chart rows, legend, and caps', () => {
    const desc = LEAN6_DESCRIPTIONS.inspect_document;
    expect(desc).toContain('r1c1');
    expect(desc).toContain('hidden-cover');
    expect(desc).toContain('[Chart column]');
    expect(desc).toContain('[TBL]');
    expect(desc).toContain('[SZ:half-points]');
    expect(desc).toContain('sec_root');
    expect(desc).toContain('limit (1-200, default 50)');
  });

  it('search description discloses chart blind spot, caps, and realistic ids', () => {
    const desc = LEAN6_DESCRIPTIONS.search_content;
    expect(desc).toContain('NOT matched by search');
    expect(desc).toContain('inspect TOON chart rows');
    expect(desc).toContain('first 50 hits');
    expect(desc).toContain('3F2A9C1B');
    expect(desc).not.toContain('p_001');
  });
});
