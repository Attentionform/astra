import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  Chart2DManipulator,
  DocumentAST,
  DocxParser,
  ToonGenerator,
  executeChart,
  is3DChartType,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

function makeMinimalAst(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body><w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:body>
    </w:document>`,
    'text/xml'
  );
  const bodyP = doc.getElementsByTagName('w:p')[0];
  return {
    metadata: {},
    sections: [{
      id: 'sec_root',
      heading: null,
      nodes: [{
        id: 'p_001',
        type: 'paragraph',
        xml: bodyP,
        metadata: {
          text: 'anchor', hasCitation: false, hasComment: false, hasBookmark: false,
          hasTrackedChanges: false, hasEquation: false, hasField: false,
          hasImage: false, hasHyperlink: false,
        },
        version: 1,
        checksum: 'chk_1',
      }],
      pageSetup: null,
      version: 1,
      checksum: 'sec_1',
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1,
  };
}

const DEMO_2D = {
  title: 'Quarterly Revenue',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'ARR', values: [10, 25, 50, 100], color: '10B981' },
    { name: 'Headcount', values: [15, 30, 60, 120], color: '6366F1' },
  ],
};

function chartXmlFromBuffer(buf: Buffer, name: string): string {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  if (!entry) throw new Error(`${name} missing from rezipped docx`);
  return entry.getData().toString('utf8');
}

describe('Chart2DManipulator Unit Suite (P2: 2D default)', () => {
  it('column emits clustered c:barChart with barDir col, gapWidth 150, overlap -30', async () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({ chartType: 'column', ...DEMO_2D });
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:barDir val="col"/>');
    expect(xml).toContain('<c:grouping val="clustered"/>');
    expect(xml).toContain('<c:gapWidth val="150"/>');
    expect(xml).toContain('<c:overlap val="-30"/>');
    expect(xml).not.toContain('3DChart');
    expect(xml).not.toContain('<c:view3D>');

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({ chartType: 'column', ...DEMO_2D });
    const buf = parser.rezip();
    expect(chartXmlFromBuffer(buf, 'chart1.xml')).toContain('<c:barDir val="col"/>');
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('bar emits clustered c:barChart with barDir bar, gapWidth 120', async () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({ chartType: 'bar', ...DEMO_2D });
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:barDir val="bar"/>');
    expect(xml).toContain('<c:grouping val="clustered"/>');
    expect(xml).toContain('<c:gapWidth val="120"/>');
    expect(xml).not.toContain('3DChart');

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({ chartType: 'bar', ...DEMO_2D });
    const buf = parser.rezip();
    expect(chartXmlFromBuffer(buf, 'chart1.xml')).toContain('<c:barDir val="bar"/>');
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('line emits c:lineChart with standard grouping, series stroke + markers', async () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({ chartType: 'line', ...DEMO_2D });
    expect(xml).toContain('<c:lineChart>');
    expect(xml).toContain('<c:grouping val="standard"/>');
    expect(xml).toContain('<c:marker>');
    expect(xml).toContain('<c:symbol val="circle"/>');
    expect(xml).toContain('<a:ln w="25400">');
    expect(xml).toContain('<c:smooth val="0"/>');

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({ chartType: 'line', ...DEMO_2D });
    const buf = parser.rezip();
    expect(chartXmlFromBuffer(buf, 'chart1.xml')).toContain('<c:lineChart>');
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('pie emits c:pieChart with varyColors 1, per-slice dPt fills, firstSliceAng 0', async () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({
      chartType: 'pie',
      title: 'Market Share',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'Share', values: [50, 30, 20], color: '0066CC' }],
    });
    expect(xml).toContain('<c:pieChart>');
    expect(xml).toContain('<c:varyColors val="1"/>');
    expect(xml).toContain('<c:dPt><c:idx val="0"/>');
    expect(xml).toContain('<c:dPt><c:idx val="1"/>');
    expect(xml).toContain('<c:firstSliceAng val="0"/>');
    expect(xml).not.toContain('<c:catAx>');

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({
      chartType: 'pie',
      title: 'Market Share',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'Share', values: [50, 30, 20], color: '0066CC' }],
    });
    const buf = parser.rezip();
    expect(chartXmlFromBuffer(buf, 'chart1.xml')).toContain('<c:pieChart>');
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('area emits c:areaChart with standard grouping', async () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({ chartType: 'area', ...DEMO_2D });
    expect(xml).toContain('<c:areaChart>');
    expect(xml).toContain('<c:grouping val="standard"/>');
    expect(xml).toContain('<c:axId val="1001"/>');

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({ chartType: 'area', ...DEMO_2D });
    const buf = parser.rezip();
    expect(chartXmlFromBuffer(buf, 'chart1.xml')).toContain('<c:areaChart>');
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('default chart type is column (omitted chartType routes to 2D)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({ action: 'chart', ...DEMO_2D } as never, ast, parser);
    const buf = parser.rezip();
    const chartXml = chartXmlFromBuffer(buf, 'chart1.xml');
    expect(chartXml).toContain('<c:barChart>');
    expect(chartXml).toContain('<c:barDir val="col"/>');
    expect(chartXml).not.toContain('3DChart');
    const node = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
    expect(node.type).toBe('chart');
    expect(node.metadata.chartType).toBe('column');
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
  });

  it('3D opt-ins still route to 3D paths and validate', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({ action: 'chart', chartType: 'column3D', ...DEMO_2D } as never, ast, parser);
    executeChart({ action: 'chart', chartType: 'bar3D', ...DEMO_2D } as never, ast, parser);
    executeChart({ action: 'chart', chartType: 'line3D', ...DEMO_2D } as never, ast, parser);
    executeChart({ action: 'chart', chartType: 'pie3D', ...DEMO_2D } as never, ast, parser);
    const buf = parser.rezip();
    const zip = new AdmZip(buf);
    const allCharts = (zip.getEntries() || [])
      .filter((e: { entryName: string }) => /^word\/charts\/chart\d+\.xml$/.test(e.entryName))
      .map((e: { getData: () => Buffer }) => e.getData().toString('utf8'))
      .join('\n');
    expect(allCharts).toContain('<c:bar3DChart>');
    expect(allCharts).toContain('<c:barDir val="col"/>');
    expect(allCharts).toContain('<c:barDir val="bar"/>');
    expect(allCharts).toContain('<c:line3DChart>');
    expect(allCharts).toContain('<c:pie3DChart>');
    const chartNodes = ast.sections[0].nodes.filter((n) => n.type === 'chart3d');
    expect(chartNodes.length).toBe(4);
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });

  it('preserves palette and series fidelity (colors, categories, values, workbook)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({ chartType: 'column', ...DEMO_2D });
    const buf = parser.rezip();
    const chartXml = chartXmlFromBuffer(buf, 'chart1.xml');
    // Series colors land as DrawingML fills.
    expect(chartXml).toContain('<a:srgbClr val="10B981"/>');
    expect(chartXml).toContain('<a:srgbClr val="6366F1"/>');
    // Categories and cached values round-trip.
    expect(chartXml).toContain('<c:v>Q1</c:v>');
    expect(chartXml).toContain('<c:v>100</c:v>');
    expect(chartXml).toContain('<c:v>120</c:v>');
    // Embedded workbook mirrors series names and values.
    const zip = new AdmZip(buf);
    const wbEntry = zip.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx');
    expect(wbEntry).not.toBeNull();
    const wbZip = new AdmZip(wbEntry!.getData());
    const sheetXml = wbZip.getEntry('xl/worksheets/sheet1.xml')?.getData().toString('utf8') ?? '';
    expect(sheetXml).toContain('ARR');
    expect(sheetXml).toContain('Headcount');
    expect(sheetXml).toContain('100');
    expect(sheetXml).toContain('120');
  });

  it('2D drawing embed carries P1 centered auto w:pPr', () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generateDrawingEmbedXml('rIdC1', 1, 'Chart 1', 5486400, 2800000);
    expect(xml).toContain('<w:pPr>');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toContain('w:lineRule="auto"');
    expect(xml).toContain('<wp:extent cx="5486400" cy="2800000"/>');
  });

  it('TOON chart tags handle new 2D type names', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({ action: 'chart', chartType: 'column', ...DEMO_2D } as never, ast, parser);
    executeChart({ action: 'chart', chartType: 'pie', title: 'P', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] } as never, ast, parser);
    const gen = new ToonGenerator();
    const table = gen.toon(ast, { fidelity: 'standard' });
    const flags = (table.rows || []).map((r) => String((r as { flags?: unknown }).flags || '')).join(' | ');
    expect(flags).toMatch(/\[CH:column\]/);
    expect(flags).toMatch(/\[CH:pie\]/);
  });

  it('reload round-trip: 2D column stays chart/column, 3D bar stays chart3d/bar3D (TOON stable)', async () => {
    // 2D default column: insert → rezip → fresh parse.
    const parser2d = DocxParser.createNewDocument('blank');
    const ast2d = await parser2d.getAST();
    executeChart({ action: 'chart', ...DEMO_2D } as never, ast2d, parser2d);
    const buf2d = parser2d.rezip();
    const fresh2d = new DocxParser(buf2d);
    const reloaded2d = await fresh2d.getAST();
    const node2d = reloaded2d.sections[0].nodes.find((n) => n.type === 'chart');
    expect(node2d).toBeDefined();
    expect(node2d!.metadata.chartType).toBe('column');
    expect(node2d!.metadata.chartData?.chartType).toBe('column');
    expect(reloaded2d.sections[0].nodes.find((n) => n.type === 'chart3d')).toBeUndefined();
    const sum2d = fresh2d.getStructuralSummary();
    const sumNode2d = sum2d.nodes.find((n: any) => n.type === 'chart');
    expect(sumNode2d).toBeDefined();
    expect(sum2d.nodes.find((n: any) => n.type === 'chart3d')).toBeUndefined();
    // TOON flag stable across reload.
    const gen = new ToonGenerator();
    const flagsBefore = (gen.toon(ast2d, { fidelity: 'standard' }).rows || [])
      .map((r) => String((r as { flags?: unknown }).flags || '')).join(' | ');
    const flagsAfter = (gen.toon(reloaded2d, { fidelity: 'standard' }).rows || [])
      .map((r) => String((r as { flags?: unknown }).flags || '')).join(' | ');
    expect(flagsBefore).toMatch(/\[CH:column\]/);
    expect(flagsAfter).toMatch(/\[CH:column\]/);
    expect(flagsAfter).not.toMatch(/bar3D/);

    // 3D control: bar3D survives the same round-trip untouched.
    const parser3d = DocxParser.createNewDocument('blank');
    const ast3d = await parser3d.getAST();
    executeChart({ action: 'chart', chartType: 'bar3D', ...DEMO_2D } as never, ast3d, parser3d);
    const buf3d = parser3d.rezip();
    const fresh3d = new DocxParser(buf3d);
    const reloaded3d = await fresh3d.getAST();
    const node3d = reloaded3d.sections[0].nodes.find((n) => n.type === 'chart3d');
    expect(node3d).toBeDefined();
    expect(node3d!.metadata.chartType).toBe('bar3D');
    expect(node3d!.metadata.chartData?.chartType).toBe('bar3D');
    const sum3d = fresh3d.getStructuralSummary();
    expect(sum3d.nodes.find((n: any) => n.type === 'chart3d')).toBeDefined();
    const flags3d = (gen.toon(reloaded3d, { fidelity: 'standard' }).rows || [])
      .map((r) => String((r as { flags?: unknown }).flags || '')).join(' | ');
    expect(flags3d).toMatch(/\[CH:bar3D\]/);
  });

  it('adversarial title: 2D column titled "My 3DChart Report" reloads as chart/column (not chart3d)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({
      action: 'chart',
      chartType: 'column',
      title: 'My 3DChart Report',
      categories: DEMO_2D.categories,
      series: DEMO_2D.series,
    } as never, ast, parser);
    const buf = parser.rezip();
    // Raw part is 2D markup despite the 3D-looking title text.
    const rawXml = chartXmlFromBuffer(buf, 'chart1.xml');
    expect(rawXml).toContain('<c:barChart>');
    expect(rawXml).toContain('My 3DChart Report');
    const fresh = new DocxParser(buf);
    const reloaded = await fresh.getAST();
    const node = reloaded.sections[0].nodes.find((n) => n.type === 'chart');
    expect(node).toBeDefined();
    expect(node!.metadata.chartType).toBe('column');
    expect(node!.metadata.chartData?.chartType).toBe('column');
    expect(reloaded.sections[0].nodes.find((n) => n.type === 'chart3d')).toBeUndefined();
    const sum = fresh.getStructuralSummary();
    expect(sum.nodes.find((n: any) => n.type === 'chart')).toBeDefined();
    expect(sum.nodes.find((n: any) => n.type === 'chart3d')).toBeUndefined();
  });

  it('is3DChartType splits 2D default from 3D opt-ins; unknown types still throw E_VALIDATION', () => {
    for (const t of ['column', 'bar', 'line', 'pie', 'area']) expect(is3DChartType(t)).toBe(false);
    for (const t of ['column3D', 'bar3D', 'line3D', 'pie3D']) expect(is3DChartType(t)).toBe(true);
    const ast = makeMinimalAst();
    try {
      executeChart({ action: 'chart', chartType: 'surface3D', title: 'Bad' } as never, ast);
      expect.unreachable('Should have thrown');
    } catch (err: unknown) {
      const e = err as Error & { code?: string; remedy?: string };
      expect(e.message).toContain('Unsupported chart type "surface3D"');
      expect(e.code).toBe('E_VALIDATION');
      expect(e.remedy).toContain('Use one of the supported chart types');
    }
    const m = new Chart2DManipulator(ast);
    expect(() => m.generate2DChartXml({ chartType: 'column3D' as never })).toThrow(/Unsupported chart type/);
  });
});

describe('Default series palette (distinct fills without explicit colors)', () => {
  function seriesFills(xml: string): string[] {
    return [...xml.matchAll(/<c:ser>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
  }

  it('column assigns distinct palette fills per uncolored series', async () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({
      chartType: 'column',
      title: 'Two series',
      categories: ['A', 'B'],
      series: [
        { name: 'Planned', values: [1, 2] },
        { name: 'Completed', values: [3, 4] },
        { name: 'Delta', values: [5, 6] },
      ],
    });
    expect(seriesFills(xml)).toEqual(['0066CC', '10B981', 'F59E0B']);

    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    new Chart2DManipulator(ast, parser).insert2DChart({
      chartType: 'column',
      title: 'Two series',
      categories: ['A', 'B'],
      series: [
        { name: 'Planned', values: [1, 2] },
        { name: 'Completed', values: [3, 4] },
      ],
    });
    const buf = parser.rezip();
    expect(seriesFills(chartXmlFromBuffer(buf, 'chart1.xml'))).toEqual(['0066CC', '10B981']);
    const validation = await validateDocxBuffer(buf);
    expect(validation.ok).toBe(true);
    expect(validation.errors).toEqual([]);
  });

  it('explicit series colors always win over the palette', () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const xml = m.generate2DChartXml({
      chartType: 'bar',
      title: 'Mixed',
      categories: ['A', 'B'],
      series: [
        { name: 'Custom', values: [1, 2], color: 'FF0000' },
        { name: 'Defaulted', values: [3, 4] },
      ],
    });
    expect(seriesFills(xml)).toEqual(['FF0000', '10B981']);
  });

  it('line defaults stroke and marker per series; single-series default stays blue', () => {
    const m = new Chart2DManipulator(makeMinimalAst());
    const solo = m.generate2DChartXml({ chartType: 'line', title: 'Solo', categories: ['A'], series: [{ name: 'Only', values: [1] }] });
    expect(seriesFills(solo)).toEqual(['0066CC']);
    const duo = m.generate2DChartXml({
      chartType: 'line', title: 'Duo', categories: ['A'],
      series: [{ name: 'One', values: [1] }, { name: 'Two', values: [2] }],
    });
    expect(seriesFills(duo)).toEqual(['0066CC', '10B981']);
  });
});
