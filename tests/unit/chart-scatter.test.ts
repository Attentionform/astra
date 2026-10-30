import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  Chart2DManipulator,
  DocxParser,
  DocumentAST,
  OperationSchema,
  SUPPORTED_CHART_2D_AND_3D_TYPES,
  ToonGenerator,
  convertAstToHtml,
  detectChartKindFromXml,
  executeChart,
  renderChartText,
  summarize2DChartXml,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

/**
 * Scatter pins: numeric x categories + per-series y values in c:scatterChart,
 * value/value axis wiring, reader/detection, in-place x/y merges, refusal
 * bounds, TOON, schema and numeric-axis HTML export.
 */

const DEMO = {
  title: 'Latency vs load',
  categories: ['0', '10', '20'],
  series: [
    { name: 'P50', values: [5, 15, 25], color: '0066CC' },
    { name: 'P99', values: [12, 34, 58], color: '10B981' },
  ],
};

const CHART_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

function toDefaultNamespace(xml: string): string {
  return xml
    .replace(`xmlns:c="${CHART_NS_URI}"`, `xmlns="${CHART_NS_URI}" xmlns:c="urn:example:vendor"`)
    .replace(/<(\/?)c:/g, '<$1');
}

function toXAlias(xml: string): string {
  return xml
    .replace(`xmlns:c="${CHART_NS_URI}"`, `xmlns:x="${CHART_NS_URI}"`)
    .replace(/<(\/?)c:/g, '<$1x:');
}

function chartXmlOf(buf: Buffer, name = 'chart1.xml'): string {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  if (!entry) throw new Error(`${name} missing from rezipped docx`);
  return entry.getData().toString('utf8');
}

function chartXml(parser: DocxParser): string {
  return parser.getZipAdapter().readText('word/charts/chart1.xml');
}

function chartNode(ast: DocumentAST, id: string) {
  return ast.sections[0].nodes.find((n) => n.id === id)!;
}

async function insertScatter(id = 'scatter', opts: Record<string, unknown> = {}) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  executeChart({ action: 'chart', chartType: 'scatter', id, ...DEMO, ...opts }, ast, parser);
  const buf = parser.rezip();
  return { parser, ast, id, buf, xml: chartXmlOf(buf) };
}

function replaceAll(text: string, from: string, to: string): string {
  return text.split(from).join(to);
}

function numLitXml(values: number[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numLit>`;
}

function sparseNumLitXml(ptCount: number, points: Array<[number, number]>): string {
  const pts = points.map(([idx, value]) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`).join('');
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${ptCount}"/>${pts}</c:numLit>`;
}

function setSeriesNumList(serXml: string, tag: 'xVal' | 'yVal', values: number[]): string {
  return serXml.replace(
    new RegExp(`<c:${tag}><c:numLit>[\\s\\S]*?</c:numLit></c:${tag}>`),
    `<c:${tag}>${numLitXml(values)}</c:${tag}>`
  );
}

function mapSeries(xml: string, index: number, map: (block: string) => string): string {
  let seen = -1;
  return xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (block) => {
    seen++;
    return seen === index ? map(block) : block;
  });
}

function xValBlocks(xml: string): string[] {
  return [...xml.matchAll(/<c:xVal>[\s\S]*?<\/c:xVal>/g)].map((match) => match[0]);
}

function yValBlocks(xml: string): string[] {
  return [...xml.matchAll(/<c:yVal>[\s\S]*?<\/c:yVal>/g)].map((match) => match[0]);
}

function axisTicks(html: string): { x: string[]; y: string[] } {
  const numeric = (text: string): boolean => /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text);
  const x = [...html.matchAll(/text-anchor="middle"[^>]*>([^<]+)<\/text>/g)]
    .map((match) => match[1].trim())
    .filter(numeric);
  const y = [...html.matchAll(/text-anchor="end"[^>]*>([^<]+)<\/text>/g)]
    .map((match) => match[1].trim())
    .filter(numeric);
  return { x, y };
}

function expectDistinctAxisTicks(html: string): void {
  const axes = axisTicks(html);
  for (const ticks of [axes.x, axes.y]) {
    expect(ticks).toHaveLength(5);
    expect(new Set(ticks).size).toBe(ticks.length);
    for (const text of ticks) expect(Number.isFinite(Number(text))).toBe(true);
  }
}

function workbookSheet(parser: DocxParser): string {
  const docx = new AdmZip(parser.rezip());
  const workbook = docx.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx');
  if (!workbook) throw new Error('embedded workbook missing');
  const xlsx = new AdmZip(workbook.getData());
  return xlsx.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8');
}

/** Chart-level region of <c:chart> before <c:plotArea>, where a chart title
 * may live; axis titles sit past this bound and must never be asserted here. */
function chartLevelXml(xml: string): string {
  const chart = /<c:chart(?:\s[^>]*)?>/.exec(xml);
  if (!chart) throw new Error('chartLevelXml: no <c:chart> element found');
  const start = chart.index + chart[0].length;
  const plotArea = /<c:plotArea(?:\s[^>]*)?>/.exec(xml);
  const end = plotArea && plotArea.index > start ? plotArea.index : xml.length;
  return xml.slice(start, end);
}

describe('scatter charts', () => {
  it('writes c:scatterChart with lineMarker style, x/y numeric literals and a valAx pair', () => {
    const xml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'scatter', ...DEMO });

    expect(xml).toContain('<c:scatterChart>');
    expect(xml).toContain('</c:scatterChart>');
    expect(xml).toContain('<c:scatterStyle val="lineMarker"/>');
    expect(xml).toContain('<c:varyColors val="0"/>');
    expect(xml.match(/<c:xVal><c:numLit>/g)).toHaveLength(2);
    expect(xml.match(/<c:yVal><c:numLit>/g)).toHaveLength(2);
    expect(xml.match(/<c:marker><c:symbol val="circle"\/>/g)).toHaveLength(2);
    expect(xml).not.toContain('<c:catAx>');
    expect(xml).not.toContain('<c:val>');
    expect(xml).not.toContain('<c:cat>');
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(xml).toContain('<c:axPos val="b"/>');
    expect(xml).toContain('<c:axPos val="l"/>');
    expect(xml).toContain('<c:crossAx val="1002"/>');
    expect(xml).toContain('<c:crossAx val="1001"/>');
    expect(xml).toContain('<c:crossBetween val="midCat"/>');

    const order = [
      '<c:scatterStyle val="lineMarker"/>',
      '<c:varyColors val="0"/>',
      '<c:ser>',
      '<c:axId val="1001"/>',
      '<c:axId val="1002"/>',
      '</c:scatterChart>',
      '<c:valAx>',
    ].map((tag) => xml.indexOf(tag));
    for (let i = 1; i < order.length; i++) {
      expect(order[i]).toBeGreaterThan(order[i - 1]);
    }
    const serXml = /<c:ser>[\s\S]*?<\/c:ser>/.exec(xml)![0];
    const serOrder = [
      '<c:idx val="0"/>',
      '<c:order val="0"/>',
      '<c:tx><c:v>',
      '<c:spPr>',
      '<c:marker>',
      '<c:xVal>',
      '<c:yVal>',
      '<c:smooth val="0"/>',
    ].map((tag) => serXml.indexOf(tag));
    for (let i = 1; i < serOrder.length; i++) {
      expect(serOrder[i]).toBeGreaterThan(serOrder[i - 1]);
    }
  });

  it('inserts a valid part with detected, summarized and hydrated scatter metadata', async () => {
    const { parser, ast, id, buf, xml } = await insertScatter();
    expect(xml).toContain('<c:scatterChart>');
    expect(detectChartKindFromXml(xml)).toEqual({ is3D: false, chartType: 'scatter' });
    const summary = summarize2DChartXml(xml);
    expect(summary.chartType).toBe('scatter');
    expect(summary.categories).toEqual(DEMO.categories);
    expect(summary.series.map((s) => s.values)).toEqual([[5, 15, 25], [12, 34, 58]]);

    const node = chartNode(ast, id);
    expect(node.type).toBe('chart');
    expect(node.metadata.chartType).toBe('scatter');
    const data = node.metadata.chartData as { chartType?: string; categories?: string[] };
    expect(data.chartType).toBe('scatter');
    expect(data.categories).toEqual(DEMO.categories);

    const docx = await validateDocxBuffer(buf);
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
    const ooxml = await parser.validateOOXML();
    expect(ooxml.ok).toBe(true);
    expect(ooxml.errors).toEqual([]);
  });

  it('surfaces scatter through TOON for read-to-update round-trips', async () => {
    const { ast, id } = await insertScatter();
    const node = chartNode(ast, id);
    expect(renderChartText(node)).toBe('[Chart scatter] Latency vs load | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58)');
    const row = new ToonGenerator().toon(ast, { fidelity: 'standard' }).rows.find((r) => r.id === id)!;
    expect(row.text).toContain('[Chart scatter]');
    expect(row.flags).toContain('[CH:scatter]');
    expect(row.flags).toMatch(/\[RID:rId\d+\]/);
  });

  it('round-trips default-namespace and x-aliased scatter parts', async () => {
    for (const spec of [
      { name: 'default', transform: toDefaultNamespace },
      { name: 'xalias', transform: toXAlias },
      {
        name: 'numeric',
        transform: (xml: string) => toDefaultNamespace(xml)
          .replace(`xmlns="${CHART_NS_URI}"`, 'xmlns="http://schemas.openxmlformats.org/drawingml/2006/char&#116;"'),
      },
    ]) {
      const base = await insertScatter(`ns_${spec.name}`);
      const aliased = spec.transform(chartXml(base.parser));
      base.parser.getZipAdapter().writeText('word/charts/chart1.xml', aliased);
      expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);
      expect(detectChartKindFromXml(aliased)).toEqual({ is3D: false, chartType: 'scatter' });
      const summary = summarize2DChartXml(aliased);
      expect(summary.chartType).toBe('scatter');
      expect(summary.title).toBe(DEMO.title);
      expect(summary.categories).toEqual(DEMO.categories);
      expect(summary.series[0]?.values).toEqual([5, 15, 25]);

      const reopened = new DocxParser(base.parser.rezip());
      const ast = await reopened.getAST();
      const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
      expect(node.metadata.chartType).toBe('scatter');
      const data = node.metadata.chartData as { title?: string; categories?: string[]; series?: Array<{ values?: number[] }> };
      expect(data.title).toBe(DEMO.title);
      expect(data.categories).toEqual(DEMO.categories);
      expect(data.series?.[0]?.values).toEqual([5, 15, 25]);
      expect(renderChartText(node)).toContain('[Chart scatter]');

      const html = convertAstToHtml(ast, `ns_${spec.name}`);
      expect(html).toContain('<polyline');
      expect(html).toContain('<title>P50: (0, 5)</title>');
      expect(html).toContain('<title>P99: (20, 58)</title>');

      const before = reopened.getZipAdapter().readText('word/charts/chart1.xml');
      executeChart({
        action: 'chart', operation: 'update', target: node.id,
        series: [{ name: 'P50', values: [7, 18, 30] }, { name: 'P99', values: [12, 34, 58] }],
      }, ast, reopened);
      const after = reopened.getZipAdapter().readText('word/charts/chart1.xml');
      const vTag = spec.name === 'xalias' ? 'x:v' : 'v';
      const expected = before
        .replace(`<${vTag}>5</${vTag}>`, `<${vTag}>7</${vTag}>`)
        .replace(`<${vTag}>15</${vTag}>`, `<${vTag}>18</${vTag}>`)
        .replace(`<${vTag}>25</${vTag}>`, `<${vTag}>30</${vTag}>`);
      expect(after).toBe(expected);
      expect(after).toMatch(spec.name === 'xalias' ? /<x:scatterChart[\s>]/ : /<scatterChart>/);
      expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
    }
  });

  it('merges values in place preserving authored scatter styling and bumping the node version', async () => {
    const { parser, ast, id } = await insertScatter('styled');
    const foreign = replaceAll(
      replaceAll(chartXml(parser), '<c:scatterStyle val="lineMarker"/>', '<c:scatterStyle val="marker"/>'),
      '<c:symbol val="circle"/>',
      '<c:symbol val="diamond"/>'
    ).replace('<c:legendPos val="r"/>', '<c:legendPos val="t"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', foreign);
    const version = chartNode(ast, id).version;

    executeChart({
      action: 'chart', operation: 'update', target: id,
      series: [{ name: 'P50', values: [7, 18, 30] }, { name: 'P99', values: [12, 34, 58] }],
    }, ast, parser);

    const after = chartXml(parser);
    const expected = foreign
      .replace('<c:v>5</c:v>', '<c:v>7</c:v>')
      .replace('<c:v>15</c:v>', '<c:v>18</c:v>')
      .replace('<c:v>25</c:v>', '<c:v>30</c:v>');
    expect(after).toBe(expected);
    expect(after).toContain('<c:scatterStyle val="marker"/>');
    expect(after.match(/<c:symbol val="diamond"\/>/g)).toHaveLength(2);
    expect(after).toContain('<c:legendPos val="t"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe('scatter');
    expect(chartNode(ast, id).version).toBe(version + 1);
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('patches every c:xVal list on a categories-only update and keeps y values', async () => {
    const { parser, ast, id } = await insertScatter('x_update');
    const before = chartXml(parser);
    executeChart({
      action: 'chart', operation: 'update', target: id, categories: ['0', '10', '30'],
    }, ast, parser);
    const after = chartXml(parser);
    expect(after).toBe(replaceAll(before, '<c:v>20</c:v>', '<c:v>30</c:v>'));
    expect(after.match(/<c:v>30<\/c:v>/g)).toHaveLength(2);
    expect(after).toContain('<c:v>58</c:v>');
    expect(chartNode(ast, id).version).toBe(2);
    expect((nodeChartData(ast, id)).categories).toEqual(['0', '10', '30']);
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('regenerates scatter on series-count and chartType changes', async () => {
    const count = await insertScatter('count');
    executeChart({
      action: 'chart', operation: 'update', target: 'count',
      series: [{ name: 'P50', values: [7, 18, 30] }],
    }, count.ast, count.parser);
    const oneSeries = chartXml(count.parser);
    expect(oneSeries).toContain('<c:scatterChart>');
    expect(oneSeries.match(/<c:ser>/g)).toHaveLength(1);
    expect(oneSeries.match(/<c:xVal>/g)).toHaveLength(1);
    expect(chartNode(count.ast, 'count').metadata.chartType).toBe('scatter');
    expect((await validateDocxBuffer(count.parser.rezip())).ok).toBe(true);

    const shape = await insertScatter('shape');
    executeChart({ action: 'chart', operation: 'update', target: 'shape', chartType: 'column' }, shape.ast, shape.parser);
    const column = chartXml(shape.parser);
    expect(column).toContain('<c:barChart>');
    expect(column).toContain('<c:catAx>');
    expect(column).not.toContain('<c:scatterChart>');
    expect(chartNode(shape.ast, 'shape').metadata.chartType).toBe('column');
    expect((await validateDocxBuffer(shape.parser.rezip())).ok).toBe(true);
  });

  it('refuses non-numeric, empty and mismatched x values without mutating', async () => {
    const { parser, ast, id } = await insertScatter('bounds');
    const before = chartXml(parser);
    const version = chartNode(ast, id).version;
    const cases = [
      { name: 'non-numeric', categories: ['a', '10', '20'], series: DEMO.series, match: /not numeric/, },
      { name: 'mismatch', categories: ['0', '10'], series: [{ name: 'P50', values: [5, 15, 25] }], match: /same length/, },
      { name: 'empty', categories: [], series: DEMO.series, match: /non-empty/, },
    ];
    for (const testCase of cases) {
      let err: (Error & { code?: string; remedy?: string }) | undefined;
      try {
        executeChart({
          action: 'chart', operation: 'update', target: id,
          categories: testCase.categories, series: testCase.series,
        }, ast, parser);
      } catch (caught) {
        err = caught as Error & { code?: string; remedy?: string };
      }
      expect(err?.message).toMatch(testCase.match);
      expect(err?.code).toBe('E_VALIDATION');
      expect(err?.remedy).toContain('numeric x values');
      expect(chartXml(parser)).toBe(before);
      expect(chartNode(ast, id).version).toBe(version);
    }

    const blank = DocxParser.createNewDocument('blank');
    const blankAst = await blank.getAST();
    expect(() => executeChart({
      action: 'chart', chartType: 'scatter', id: 'bad',
      categories: ['1', '2'], series: [{ name: 'S', values: [1] }],
    }, blankAst, blank)).toThrow(/same length/);
    expect(blank.getZipAdapter().hasFile('word/charts/chart1.xml')).toBe(false);
    expect(blankAst.sections.flatMap((s) => s.nodes).some((n) => n.type === 'chart')).toBe(false);

    expect(() => new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'scatter', categories: [''], series: [{ name: 'S', values: [1] }] }))
      .toThrow(/not numeric/);
  });

  it('refuses multi-plot and same-tag duplicate scatter parts before mutation', async () => {
    const base = await insertScatter('combo');
    const singleXml = chartXml(base.parser);
    const barXml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'bar', title: 'Other', categories: ['Q1', 'Q2'], series: [{ name: 'Other', values: [3, 4] }] });
    const barPlot = /<c:barChart>[\s\S]*?<\/c:barChart>/.exec(barXml)![0];
    const combo = singleXml.replace('</c:scatterChart>', `</c:scatterChart>${barPlot}`);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', combo);
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);

    const scatterBlock = /<c:scatterChart>[\s\S]*?<\/c:scatterChart>/.exec(singleXml)![0];
    const fixtures = [
      { xml: combo, tags: 'barChart, scatterChart' },
      { xml: singleXml.replace('</c:plotArea>', `${scatterBlock}</c:plotArea>`), tags: 'scatterChart, scatterChart' },
    ];
    for (const fixture of fixtures) {
      base.parser.getZipAdapter().writeText('word/charts/chart1.xml', fixture.xml);
      for (const series of [DEMO.series, [DEMO.series[0]]]) {
        const version = chartNode(base.ast, base.id).version;
        let err: (Error & { code?: string; remedy?: string }) | undefined;
        try {
          executeChart({ action: 'chart', operation: 'update', target: base.id, series }, base.ast, base.parser);
        } catch (caught) {
          err = caught as Error & { code?: string; remedy?: string };
        }
        expect(err?.message).toContain('multi-plot chart part');
        expect(err?.message).toContain(fixture.tags);
        expect(err?.message).toContain('word/charts/chart1.xml');
        expect(err?.remedy).toContain('single-plot');
        expect(chartXml(base.parser)).toBe(fixture.xml);
        expect(chartNode(base.ast, base.id).version).toBe(version);
      }
    }
  });

  it('refuses non-dense x indexes with a remedy and no mutation', async () => {
    const { parser, ast, id } = await insertScatter('dense');
    const before = chartXml(parser).replace(/(<c:xVal>[\s\S]*?<\/c:xVal>)/, (block) => block.replace('idx="1"', 'idx="5"'));
    parser.getZipAdapter().writeText('word/charts/chart1.xml', before);
    const sparse = new DocxParser(parser.rezip());
    const sparseAst = await sparse.getAST();
    const sparseNode = sparseAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const sparseText = renderChartText(sparseNode);
    expect(sparseText).toContain('P50: (0, 5),(20, 25); 1 y value without x');
    expect(sparseText).not.toContain('(20, 15)');
    const sparseHtml = convertAstToHtml(sparseAst, 'sparse');
    expect(sparseHtml).toContain('<title>P50: (20, 25)</title>');
    expect(sparseHtml).not.toContain('<title>P50: (20, 15)</title>');
    const version = chartNode(ast, id).version;
    let err: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({ action: 'chart', operation: 'update', target: id, categories: ['0', '10', '30'] }, ast, parser);
    } catch (caught) {
      err = caught as Error & { code?: string; remedy?: string };
    }
    expect(err?.message).toContain('non-dense point indexes');
    expect(err?.code).toBe('E_VALIDATION');
    expect(err?.remedy).toContain('inspect_document');
    expect(chartXml(parser)).toBe(before);
    expect(chartNode(ast, id).version).toBe(version);
  });

  it('operation schema accepts scatter in the chart and legacy insert shapes', () => {
    expect(SUPPORTED_CHART_2D_AND_3D_TYPES).toContain('scatter');
    expect(OperationSchema.safeParse({
      action: 'chart', operation: 'insert', anchor: 'p_001', chartType: 'scatter',
      categories: ['0', '1'], series: [{ name: 'S', values: [1, 2] }],
    }).success).toBe(true);
    expect(OperationSchema.safeParse({
      action: 'insert', type: 'chart', anchor: 'p_001', position: 'after', chartType: 'scatter',
      categories: ['0', '1'], series: [{ name: 'S', values: [1, 2] }],
    }).success).toBe(true);
    expect(OperationSchema.safeParse({
      action: 'chart', operation: 'insert', anchor: 'p_001', chartType: 'bubble',
      categories: ['0', '1'], series: [{ name: 'S', values: [1, 2] }],
    }).success).toBe(false);
  });

  it('HTML export draws points and lines on numeric scaled axes', async () => {
    const { ast } = await insertScatter('html');
    const html = convertAstToHtml(ast, 'html');
    expect(html.match(/<circle /g)).toHaveLength(6);
    expect(html.match(/<polyline /g)).toHaveLength(2);
    expect(html).toContain('<title>P50: (0, 5)</title>');
    expect(html).toContain('<title>P50: (20, 25)</title>');
    expect(html).toContain('<title>P99: (20, 58)</title>');
    expectDistinctAxisTicks(html);

    expect(html).toContain('P50');
    expect(html).toContain('P99');
  });

  it('represents and preserves legal per-series x lists across TOON, HTML and updates', async () => {
    const base = await insertScatter('per_series');
    const divergent = mapSeries(base.xml, 1, (block) =>
      setSeriesNumList(setSeriesNumList(block, 'xVal', [0, 10, 20, 30]), 'yVal', [12, 34, 58, 80])
    );
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', divergent);

    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    const data = node.metadata.chartData as {
      categories: string[];
      series: Array<{ name: string; values: number[]; xValues?: number[] }>;
      synthetic?: boolean;
    };
    expect(data.synthetic).toBeUndefined();
    expect(data.categories).toEqual(['0', '10', '20']);
    expect(data.series[0].xValues).toEqual([0, 10, 20]);
    expect(data.series[1].xValues).toEqual([0, 10, 20, 30]);
    expect(renderChartText(node)).toBe(
      '[Chart scatter] Latency vs load | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58),(30, 80)'
    );
    const row = new ToonGenerator().toon(ast, { fidelity: 'standard' }).rows.find((r) => r.id === node.id)!;
    expect(row.text).toContain('P99: (0, 12),(10, 34),(20, 58),(30, 80)');

    const html = convertAstToHtml(ast, 'per_series');
    expect(html).toContain('<title>P50: (20, 25)</title>');
    expect(html).toContain('<title>P99: (30, 80)</title>');
    expect(html).not.toContain('P99: (20, 80)');

    const beforeTitle = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    const beforeSheet = workbookSheet(reopened);
    executeChart({ action: 'chart', operation: 'update', target: node.id, title: 'Retitled' }, ast, reopened);
    const afterTitle = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(afterTitle.replace(/<c:title>[\s\S]*?<\/c:title>/, '')).toBe(
      beforeTitle.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    expect(xValBlocks(afterTitle)).toEqual(xValBlocks(beforeTitle));
    expect((node.metadata.chartData as { series: Array<{ xValues?: number[] }> }).series[1].xValues).toEqual([0, 10, 20, 30]);
    expect(workbookSheet(reopened)).toBe(beforeSheet);
    expect(node.version).toBe(2);

    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'P50', values: [6, 16, 26] }, { name: 'P99', values: [13, 35, 59, 81] }],
    }, ast, reopened);
    const afterValues = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(yValBlocks(afterValues)).toEqual([
      numLitXml([6, 16, 26]),
      numLitXml([13, 35, 59, 81]),
    ].map((block) => block.replace('<c:numLit>', '<c:yVal><c:numLit>').replace('</c:numLit>', '</c:numLit></c:yVal>')));
    expect(xValBlocks(afterValues)).toEqual(xValBlocks(beforeTitle));
    expect(node.version).toBe(3);

    executeChart({ action: 'chart', operation: 'update', target: node.id, categories: ['2', '12', '22'] }, ast, reopened);
    const afterCategories = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(xValBlocks(afterCategories)).toEqual([
      `<c:xVal>${numLitXml([2, 12, 22])}</c:xVal>`,
      xValBlocks(beforeTitle)[1],
    ]);
    expect(node.version).toBe(4);

    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [
        { name: 'P50', values: [6, 16, 26], xValues: [2, 12, 22] },
        { name: 'P99', values: [13, 35, 59, 81], xValues: [3, 13, 23, 33] },
      ],
    }, ast, reopened);
    const afterX = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(xValBlocks(afterX)).toEqual(
      [numLitXml([2, 12, 22]), numLitXml([3, 13, 23, 33])].map((block) => `<c:xVal>${block}</c:xVal>`)
    );
    expect(yValBlocks(afterX)).toEqual(yValBlocks(afterValues));
    expect((node.metadata.chartData as { categories: string[] }).categories).toEqual(['2', '12', '22']);
    expect(node.version).toBe(5);
    expect(convertAstToHtml(ast, 'per_series')).toContain('<title>P99: (33, 81)</title>');
    expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
  });

  it('keeps bubble parts synthetic and unhydrated', async () => {
    const base = await insertScatter('bubble_base');
    const scatterBlock = /<c:scatterChart>[\s\S]*?<\/c:scatterChart>/.exec(base.xml)![0];
    const bubbleBlock = `<c:bubbleChart>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:v>Bubbles</c:v></c:tx>
          <c:xVal>${numLitXml([1, 2])}</c:xVal>
          <c:yVal>${numLitXml([3, 4])}</c:yVal>
          <c:bubbleSize>${numLitXml([5, 6])}</c:bubbleSize>
        </c:ser>
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:bubbleChart>`;
    const bubbleXml = base.xml.replace(scatterBlock, bubbleBlock);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', bubbleXml);

    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    const data = node.metadata.chartData as {
      categories: string[];
      series: Array<{ values: number[] }>;
      synthetic?: boolean;
    };
    expect(data.synthetic).toBe(true);
    expect(data.categories).toEqual([]);
    expect(data.series[0].values).toEqual([]);
    expect(renderChartText(node)).toBe('[Chart ?]');
    const summary = summarize2DChartXml(bubbleXml);
    expect(summary.chartType).not.toBe('scatter');
    expect(summary.categories).toEqual([]);
    expect(summary.series[0].values).toEqual([]);
  });

  it('rewrites numRef scatter formulas to the regenerated workbook on every merge', async () => {
    const base = await insertScatter('formula', {
      series: [{ name: 'S', values: [5, 15, 25], color: '0066CC' }],
    });
    const cache = (values: number[]): string =>
      `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
      values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
      '</c:numCache>';
    const ref = (formula: string, values: number[]): string => `<c:numRef><c:f>${formula}</c:f>${cache(values)}</c:numRef>`;
    const foreign = base.xml
      .replace(/<c:xVal><c:numLit>[\s\S]*?<\/c:numLit><\/c:xVal>/, `<c:xVal>${ref('Sheet1!$B$2:$B$4', [0, 10, 20])}</c:xVal>`)
      .replace(/<c:yVal><c:numLit>[\s\S]*?<\/c:numLit><\/c:yVal>/, `<c:yVal>${ref('Sheet1!$C$2:$C$4', [5, 15, 25])}</c:yVal>`);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', foreign);
    const titleOnlyXml = chartXml(base.parser);
    const titleOnlySheet = workbookSheet(base.parser);
    executeChart({ action: 'chart', operation: 'update', target: base.id, title: 'Retitled' }, base.ast, base.parser);
    expect(chartXml(base.parser)).toContain('<c:f>Sheet1!$B$2:$B$4</c:f>');
    expect(chartXml(base.parser)).toContain('<c:f>Sheet1!$C$2:$C$4</c:f>');
    expect(chartXml(base.parser).replace(/<c:title>[\s\S]*?<\/c:title>/, '')).toBe(
      titleOnlyXml.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    expect(workbookSheet(base.parser)).toBe(titleOnlySheet);

    executeChart({
      action: 'chart', operation: 'update', target: base.id,
      series: [{ name: 'S', values: [7, 18, 30] }],
    }, base.ast, base.parser);
    let xml = chartXml(base.parser);
    expect(xml).toContain('<c:f>Sheet1!$A$2:$A$4</c:f>');
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$4</c:f>');
    let sheet = workbookSheet(base.parser);
    expect(sheet).toContain('<c r="A2"><v>0</v></c>');
    expect(sheet).toContain('<c r="B2"><v>7</v></c>');
    expect(sheet).toContain('<c r="A4"><v>20</v></c>');
    expect(sheet).toContain('<c r="B4"><v>30</v></c>');

    executeChart({
      action: 'chart', operation: 'update', target: base.id,
      categories: ['0', '10', '20', '30'],
      series: [{ name: 'S', values: [7, 18, 30, 44] }],
    }, base.ast, base.parser);
    xml = chartXml(base.parser);
    expect(xml).toContain('<c:f>Sheet1!$A$2:$A$5</c:f>');
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$5</c:f>');
    sheet = workbookSheet(base.parser);
    expect(sheet).toContain('<c r="A5"><v>30</v></c>');
    expect(sheet).toContain('<c r="B5"><v>44</v></c>');
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);
  });

  it('formats small, large and extreme finite scatter ranges with finite geometry', async () => {
    const small = await insertScatter('small_range', {
      categories: ['0.00001', '0.00002', '0.00003'],
      series: [{ name: 'S', values: [0.000001, 0.000002, 0.000003] }],
    });
    const smallHtml = convertAstToHtml(small.ast, 'small_range');
    expect(smallHtml).toContain('<title>S: (1e-5, 1e-6)</title>');
    expect(smallHtml).not.toContain('(0, 0)');
    expectDistinctAxisTicks(smallHtml);

    const large = await insertScatter('large_range', {
      categories: ['1000000', '2000000', '3000000'],
      series: [{ name: 'L', values: [10000000, 20000000, 30000000] }],
    });
    const largeHtml = convertAstToHtml(large.ast, 'large_range');
    expect(largeHtml).toContain('<title>L: (1000000, 10000000)</title>');
    expect(largeHtml).not.toMatch(/NaN|Infinity/);

    const extreme = await insertScatter('extreme_range', {
      categories: ['-1e308', '1e308'],
      series: [{ name: 'E', values: [-1e308, 1e308] }],
    });
    const extremeHtml = convertAstToHtml(extreme.ast, 'extreme_range');
    expect(extremeHtml).not.toMatch(/NaN|Infinity/);
    const geometry = [...extremeHtml.matchAll(/(?:x1|y1|x2|y2|cx|cy)="([^"]+)"/g)].map((m) => Number(m[1]));
    expect(geometry.length).toBeGreaterThan(10);
    for (const value of geometry) expect(Number.isFinite(value)).toBe(true);
    expectDistinctAxisTicks(extremeHtml);
  });

  it('rejects non-finite x and y values before writing', async () => {
    const blank = DocxParser.createNewDocument('blank');
    const blankAst = await blank.getAST();
    expect(() => executeChart({
      action: 'chart', chartType: 'scatter', id: 'overflow',
      categories: ['1e999', '2'], series: [{ name: 'S', values: [1, 2] }],
    }, blankAst, blank)).toThrow(/not finite/);
    expect(() => executeChart({
      action: 'chart', chartType: 'scatter', id: 'overflow',
      categories: ['1', '2'], series: [{ name: 'S', values: [Number('1e999'), 2] }],
    }, blankAst, blank)).toThrow(/non-finite y value/);
    expect(blank.getZipAdapter().hasFile('word/charts/chart1.xml')).toBe(false);
    expect(blankAst.sections.flatMap((s) => s.nodes).some((n) => n.type === 'chart')).toBe(false);

    const base = await insertScatter('overflow_update');
    const before = chartXml(base.parser);
    const version = chartNode(base.ast, base.id).version;
    expect(() => executeChart({
      action: 'chart', operation: 'update', target: base.id,
      categories: ['1e999', '2', '3'],
    }, base.ast, base.parser)).toThrow(/not finite/);
    expect(() => executeChart({
      action: 'chart', operation: 'update', target: base.id,
      series: [{ name: 'P50', values: [5, 15, Infinity] }, { name: 'P99', values: [12, 34, 58] }],
    }, base.ast, base.parser)).toThrow(/non-numeric value|non-finite y value/);
    expect(chartXml(base.parser)).toBe(before);
    expect(chartNode(base.ast, base.id).version).toBe(version);
  });

  it('refuses x-value span/series skew on a categories update without mutating', async () => {
    const base = await insertScatter('span_skew');
    const missingX = mapSeries(base.xml, 1, (block) => block.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, ''));
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', missingX);
    const version = chartNode(base.ast, base.id).version;
    let refusal: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({ action: 'chart', operation: 'update', target: base.id, categories: ['0', '10', '30'] }, base.ast, base.parser);
    } catch (caught) {
      refusal = caught as Error & { code?: string; remedy?: string };
    }
    expect(refusal?.message).toMatch(/no <c:xVal> block|x-value lists but 2 series/);
    expect(refusal?.code).toBe('E_VALIDATION');
    expect(refusal?.remedy).toContain('c:xVal');
    expect(chartXml(base.parser)).toBe(missingX);
    expect(chartNode(base.ast, base.id).version).toBe(version);
  });

  it('patches scatter series color without touching x or y data', async () => {
    const base = await insertScatter('color_update');
    const before = chartXml(base.parser);
    executeChart({
      action: 'chart', operation: 'update', target: base.id,
      series: [
        { name: 'P50', values: [5, 15, 25], color: 'FF0000' },
        { name: 'P99', values: [12, 34, 58], color: '10B981' },
      ],
    }, base.ast, base.parser);
    const after = chartXml(base.parser);
    expect(xValBlocks(after)).toEqual(xValBlocks(before));
    expect(yValBlocks(after)).toEqual(yValBlocks(before));
    const firstSeries = /<c:ser>[\s\S]*?<\/c:ser>/.exec(after)![0];
    expect(firstSeries.match(/<a:srgbClr val="FF0000"\/>/g)).toHaveLength(2);
    expect(after).toContain('<a:srgbClr val="10B981"/>');
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);
  });

  it('grows a shared scatter point count through the merge and workbook path', async () => {
    const base = await insertScatter('growth');
    executeChart({
      action: 'chart', operation: 'update', target: base.id,
      categories: ['0', '10', '20', '30'],
      series: [
        { name: 'P50', values: [5, 15, 25, 35] },
        { name: 'P99', values: [12, 34, 58, 80] },
      ],
    }, base.ast, base.parser);
    const xml = chartXml(base.parser);
    expect(xValBlocks(xml)).toHaveLength(2);
    for (const block of xValBlocks(xml)) {
      expect(block).toContain('<c:ptCount val="4"/>');
      expect(block).toContain('<c:v>30</c:v>');
    }
    expect(yValBlocks(xml)[0]).toContain('<c:ptCount val="4"/>');
    expect(yValBlocks(xml)[1]).toContain('<c:v>80</c:v>');
    const sheet = workbookSheet(base.parser);
    expect(sheet).toContain('<c r="A5"><v>30</v></c>');
    expect(sheet).toContain('<c r="B5"><v>35</v></c>');
    expect(sheet).toContain('<c r="D5"><v>80</v></c>');
    expect((nodeChartData(base.ast, base.id)).categories).toEqual(['0', '10', '20', '30']);
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);
  });


  it('keeps mixed shared and per-series x consistent on insert and update', async () => {
    const mixed = await insertScatter('mixed_x', {
      categories: ['100', '200'],
      series: [
        { name: 'A', values: [1, 2], xValues: [0, 1] },
        { name: 'B', values: [3, 4] },
      ],
    });
    const inserted = chartNode(mixed.ast, mixed.id).metadata.chartData as {
      categories: string[];
      series: Array<{ name: string; xValues?: number[] }>;
    };
    expect(inserted.categories).toEqual(['0', '1']);
    expect(inserted.series[0].xValues).toEqual([0, 1]);
    expect(inserted.series[1].xValues).toEqual([100, 200]);
    expect(xValBlocks(mixed.xml)).toEqual([
      `<c:xVal>${numLitXml([0, 1])}</c:xVal>`,
      `<c:xVal>${numLitXml([100, 200])}</c:xVal>`,
    ]);
    expect(renderChartText(chartNode(mixed.ast, mixed.id))).toBe(
      '[Chart scatter] Latency vs load | A: (0, 1),(1, 2) | B: (100, 3),(200, 4)'
    );
    const insertHtml = convertAstToHtml(mixed.ast, 'mixed_x');
    expect(insertHtml).toContain('<title>B: (200, 4)</title>');
    expect(insertHtml).not.toContain('<title>B: (1, 4)</title>');

    const updated = await insertScatter('mixed_update');
    executeChart({
      action: 'chart', operation: 'update', target: updated.id,
      series: [
        { name: 'P50', values: [5, 15, 25], xValues: [2, 12, 22] },
        { name: 'P99', values: [12, 34, 58] },
      ],
    }, updated.ast, updated.parser);
    const reopened = new DocxParser(updated.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    const data = node.metadata.chartData as {
      categories: string[];
      series: Array<{ name: string; xValues?: number[] }>;
    };
    expect(xValBlocks(chartXml(updated.parser))).toEqual([
      `<c:xVal>${numLitXml([2, 12, 22])}</c:xVal>`,
      `<c:xVal>${numLitXml([0, 10, 20])}</c:xVal>`,
    ]);
    expect(data.categories).toEqual(['2', '12', '22']);
    expect(data.series[1].xValues).toEqual([0, 10, 20]);
    expect(renderChartText(node)).toBe(
      '[Chart scatter] Latency vs load | P50: (2, 5),(12, 15),(22, 25) | P99: (0, 12),(10, 34),(20, 58)'
    );
    const html = convertAstToHtml(ast, 'mixed_update');
    expect(html).toContain('<title>P99: (10, 34)</title>');
    expect(html).not.toContain('<title>P99: (12, 34)</title>');
  });

  it('keeps a title-less part titleless through data updates and handles an absent first x', async () => {
    const titleless = await insertScatter('titleless_data');
    titleless.parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      titleless.xml.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const reopened = new DocxParser(titleless.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'P50', values: [6, 16, 26] }, { name: 'P99', values: [13, 35, 59] }],
    }, ast, reopened);
    expect(chartXml(reopened)).not.toContain('<c:title>');
    expect(node.metadata.title).toBeUndefined();
    expect(Object.hasOwn(node.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(node.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(node)).not.toContain('Chart Analysis');
    expect(renderChartText(node)).toBe('[Chart scatter] | P50: (0, 6),(10, 16),(20, 26) | P99: (0, 13),(10, 35),(20, 59)');
    expect(convertAstToHtml(ast, 'titleless_data')).not.toContain('Chart Analysis');

    const empty = await insertScatter('empty_titled');
    const noData = empty.xml
      .replace(/<c:ser>[\s\S]*?<\/c:ser>/g, '')
      .replace(/<c:cat>[\s\S]*?<\/c:cat>/g, '');
    empty.parser.getZipAdapter().writeText('word/charts/chart1.xml', noData);
    const emptyReopened = new DocxParser(empty.parser.rezip());
    const emptyAst = await emptyReopened.getAST();
    const emptyNode = emptyAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(emptyNode)).toBe('[Chart ?]');
    executeChart({ action: 'chart', operation: 'update', target: emptyNode.id, title: 'Only Title' }, emptyAst, emptyReopened);
    expect(renderChartText(emptyNode)).toBe('[Chart ?]');

    const absentFirst = await insertScatter('absent_first');
    const noFirstX = mapSeries(absentFirst.xml, 0, (block) => block.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, ''));
    absentFirst.parser.getZipAdapter().writeText('word/charts/chart1.xml', noFirstX);
    const firstReopened = new DocxParser(absentFirst.parser.rezip());
    const firstAst = await firstReopened.getAST();
    const firstNode = firstAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(firstNode)).toContain('P50: 3 y values without x');
    expect(renderChartText(firstNode)).toContain('P99: (0, 12),(10, 34),(20, 58)');
    expect(renderChartText(firstNode)).not.toBe('[Chart ?]');
    const firstHtml = convertAstToHtml(firstAst, 'absent_first');
    expect(firstHtml).toContain('3 points omitted: no x value');
    expect(firstHtml).toContain('<title>P99: (10, 34)</title>');
    const firstSheet = workbookSheet(firstReopened);
    executeChart({ action: 'chart', operation: 'update', target: firstNode.id, title: 'Retitled' }, firstAst, firstReopened);
    expect(workbookSheet(firstReopened)).toBe(firstSheet);
    const retitledText = renderChartText(firstNode);
    expect(retitledText).toContain('[Chart scatter] Retitled');
    expect(retitledText).toContain('P99: (0, 12),(10, 34),(20, 58)');
    expect(convertAstToHtml(firstAst, 'absent_first')).toContain('<title>P99: (10, 34)</title>');

    const catOnly = await insertScatter('cat_only');
    const catBlock = '<c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strLit></c:cat>';
    const catPart = mapSeries(
      mapSeries(catOnly.xml, 0, (block) => block.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, catBlock)),
      1,
      (block) => block.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, '')
    );
    catOnly.parser.getZipAdapter().writeText('word/charts/chart1.xml', catPart);
    const catReopened = new DocxParser(catOnly.parser.rezip());
    const catAst = await catReopened.getAST();
    const catNode = catAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(catNode)).toContain('cats: Q1,Q2,Q3');
    expect(renderChartText(catNode)).not.toBe('[Chart ?]');
    const catHtml = convertAstToHtml(catAst, 'cat_only');
    expect(catHtml).toContain('x values are not numeric; scatter points not plotted');
    expect(catHtml).not.toContain('points omitted: no x value');
    executeChart({ action: 'chart', operation: 'update', target: catNode.id, title: 'Cats Retitled' }, catAst, catReopened);
    expect(renderChartText(catNode)).toContain('[Chart scatter] Cats Retitled');
    expect(renderChartText(catNode)).toContain('cats: Q1,Q2,Q3');
    const catXmlBefore = chartXml(catReopened);
    let catErr: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({ action: 'chart', operation: 'update', target: catNode.id, categories: ['1', '2', '3'] }, catAst, catReopened);
    } catch (caught) {
      catErr = caught as Error & { code?: string; remedy?: string };
    }
    expect(catErr?.code).toBe('E_VALIDATION');
    expect(catErr?.message).toContain('no <c:xVal> block');
    expect(catErr?.remedy).toContain('c:xVal');
    expect(chartXml(catReopened)).toBe(catXmlBefore);
  });

  it('keeps title-less legacy charts hydrated under the shared synthetic rule', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({
      action: 'chart', chartType: 'column', id: 'legacy_titleless', title: 'Temp',
      categories: ['Q1', 'Q2'], series: [{ name: 'S', values: [1, 2] }],
    }, ast, parser);
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      chartXml(parser).replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const node = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartData?.synthetic).toBeUndefined();
    expect(renderChartText(node)).toBe('[Chart column] | cats: Q1,Q2 | S: 1,2');
  });

  it('accepts a series-only scatter payload when every series has xValues', async () => {
    const seriesOnly = await insertScatter('series_only', {
      categories: undefined,
      series: [
        { name: 'A', values: [1, 2], xValues: [10, 20] },
        { name: 'B', values: [3, 4], xValues: [30, 40] },
      ],
    });
    expect(xValBlocks(seriesOnly.xml)).toEqual([
      `<c:xVal>${numLitXml([10, 20])}</c:xVal>`,
      `<c:xVal>${numLitXml([30, 40])}</c:xVal>`,
    ]);
    const data = chartNode(seriesOnly.ast, seriesOnly.id).metadata.chartData as {
      categories: string[];
      series: Array<{ xValues?: number[] }>;
    };
    expect(data.categories).toEqual(['10', '20']);
    expect(data.series[0].xValues).toEqual([10, 20]);
    expect(data.series[1].xValues).toEqual([30, 40]);
  });

  it('renders a title-less scatter and reports points without x instead of inventing x', async () => {
    const titleless = await insertScatter('titleless');
    titleless.parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      titleless.xml.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const reopened = new DocxParser(titleless.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartData?.synthetic).toBeUndefined();
    expect(renderChartText(node)).toContain('[Chart scatter]');
    const html = convertAstToHtml(ast, 'titleless');
    expect(html).not.toContain('Scatter data not rendered');
    expect(html.match(/<circle /g)).toHaveLength(6);
    expect(html).toContain('<title>P50: (10, 15)</title>');
    executeChart({ action: 'chart', operation: 'update', target: node.id, title: 'Now Titled' }, ast, reopened);
    expect(node.metadata.chartData?.synthetic).toBeUndefined();
    expect(renderChartText(node)).toContain('[Chart scatter] Now Titled');

    const short = await insertScatter('short_x');
    const shortXml = mapSeries(short.xml, 1, (block) =>
      block.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, `<c:xVal>${numLitXml([9])}</c:xVal>`)
    );
    short.parser.getZipAdapter().writeText('word/charts/chart1.xml', shortXml);
    const shortReopened = new DocxParser(short.parser.rezip());
    const shortAst = await shortReopened.getAST();
    const shortNode = shortAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const text = renderChartText(shortNode);
    expect(text).toContain('P99: (9, 12); 2 y values without x');
    expect(text).not.toContain('(1, 34)');
    const shortHtml = convertAstToHtml(shortAst, 'short_x');
    expect(shortHtml).toContain('2 points omitted: no x value');
    expect(shortHtml).not.toContain('<title>P99: (1, 34)</title>');

    const absent = await insertScatter('absent_x');
    const absentXml = mapSeries(absent.xml, 1, (block) => block.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, ''));
    absent.parser.getZipAdapter().writeText('word/charts/chart1.xml', absentXml);
    const absentReopened = new DocxParser(absent.parser.rezip());
    const absentAst = await absentReopened.getAST();
    const absentNode = absentAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const absentText = renderChartText(absentNode);
    expect(absentText).toContain('P99: 3 y values without x');
    expect(absentText).not.toContain('(1, 12)');
    const absentHtml = convertAstToHtml(absentAst, 'absent_x');
    expect(absentHtml).toContain('3 points omitted: no x value');
    expect(absentHtml).not.toContain('<title>P99: (1, 12)</title>');
  });

  it('formats tight timestamp and tiny ranges with distinct span-aware labels', async () => {
    const stamps = await insertScatter('stamps', {
      categories: ['1700000000', '1700000060'],
      series: [{ name: 'T', values: [1, 2] }],
    });
    const stampsHtml = convertAstToHtml(stamps.ast, 'stamps');
    expectDistinctAxisTicks(stampsHtml);
    expect(stampsHtml).toContain('<title>T: (1700000000, 1)</title>');
    expect(stampsHtml).toContain('<title>T: (1700000060, 2)</title>');
    const stampTicks = axisTicks(stampsHtml).x;
    expect(new Set(stampTicks).size).toBe(5);
    expect(stampTicks.every((text) => text.length <= 12)).toBe(true);

    const tiny = await insertScatter('tiny', {
      categories: ['1e-9', '1.0000001e-9'],
      series: [{ name: 'N', values: [1e-9, 2e-9] }],
    });
    const tinyHtml = convertAstToHtml(tiny.ast, 'tiny');
    expectDistinctAxisTicks(tinyHtml);
    expect(tinyHtml).not.toMatch(/NaN|Infinity/);

    const hugeTight = await insertScatter('huge_tight', {
      categories: ['1000000000000000', '1000000000000001'],
      series: [{ name: 'H', values: [1, 2] }],
    });
    const hugeHtml = convertAstToHtml(hugeTight.ast, 'huge_tight');
    expectDistinctAxisTicks(hugeHtml);
    const hugeTips = [...hugeHtml.matchAll(/<title>H: \(([^,]+),/g)].map((m) => m[1]);
    expect(new Set(hugeTips).size).toBe(2);

    const tinyTight = await insertScatter('tiny_tight', {
      categories: ['1.0000000001', '1.0000000002'],
      series: [{ name: 'S', values: [1, 2] }],
    });
    const tightHtml = convertAstToHtml(tinyTight.ast, 'tiny_tight');
    expectDistinctAxisTicks(tightHtml);
    const tightTips = [...tightHtml.matchAll(/<title>S: \(([^,]+),/g)].map((m) => m[1]);
    expect(new Set(tightTips).size).toBe(2);
  });

  it('keeps a legal bar+scatter combo on its pre-existing bar family when read or exported', async () => {
    const base = await insertScatter('combo_read');
    const singleXml = chartXml(base.parser);
    const barXml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'bar', title: 'Other', categories: ['Q1', 'Q2'], series: [{ name: 'Other', values: [3, 4] }] });
    const barPlot = /<c:barChart>[\s\S]*?<\/c:barChart>/.exec(barXml)![0];
    const combo = singleXml.replace('</c:scatterChart>', `</c:scatterChart>${barPlot}`);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', combo);
    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('bar');
    expect(renderChartText(node)).toContain('[Chart bar]');
    const summary = summarize2DChartXml(combo);
    expect(summary.chartType).toBe('bar');
    expect(summary.categories).toEqual(['Q1', 'Q2']);
    const html = convertAstToHtml(ast, 'combo_read');
    expect(html).not.toContain('Scatter data not rendered');
    expect(html).toMatch(/<rect /);
  });

function nodeChartData(ast: DocumentAST, id: string): { chartType?: string; categories?: string[] } {
  return chartNode(ast, id).metadata.chartData as { chartType?: string; categories?: string[] };
}

  it('regenerates scatter x and workbook when switching a column chart to scatter', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({
      action: 'chart', chartType: 'column', id: 'switch',
      categories: ['0', '10', '20'], series: [{ name: 'S', values: [1, 2, 3] }],
    }, ast, parser);
    executeChart({ action: 'chart', operation: 'update', target: 'switch', chartType: 'scatter' }, ast, parser);
    const node = ast.sections[0].nodes.find((n) => n.id === 'switch')!;
    const data = node.metadata.chartData as { series: Array<{ xValues?: number[] }> };
    expect(data.series[0].xValues).toEqual([0, 10, 20]);
    expect(xValBlocks(chartXml(parser))).toEqual([`<c:xVal>${numLitXml([0, 10, 20])}</c:xVal>`]);
    const sheet = workbookSheet(parser);
    expect(sheet).toContain('<c r="A2"><v>0</v></c>');
    expect(sheet).toContain('<c r="A4"><v>20</v></c>');
    expect(sheet).toContain('<c r="B2"><v>1</v></c>');
    expect(sheet).toContain('<c r="B4"><v>3</v></c>');
    expect(renderChartText(node)).toContain('S: (0, 1),(10, 2),(20, 3)');
  });

  it('gives a newly added series the existing shared x list', async () => {
    const base = await insertScatter('grow_series');
    executeChart({
      action: 'chart', operation: 'update', target: base.id,
      series: [
        { name: 'P50', values: [5, 15, 25] },
        { name: 'P99', values: [12, 34, 58] },
        { name: 'P999', values: [20, 40, 60] },
      ],
    }, base.ast, base.parser);
    expect(xValBlocks(chartXml(base.parser))).toHaveLength(3);
    for (const block of xValBlocks(chartXml(base.parser))) {
      expect(block).toContain('<c:v>10</c:v>');
    }
    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(node)).toContain('P999: (0, 20),(10, 40),(20, 60)');
  });

  it('keeps per-series x matched by name when series are removed, reordered or prepended', async () => {
    const divergent = {
      categories: undefined,
      series: [
        { name: 'A', values: [1, 2], xValues: [0, 10] },
        { name: 'B', values: [3, 4], xValues: [5, 15] },
      ],
    };

    const removed = await insertScatter('x_remove', divergent);
    executeChart({
      action: 'chart', operation: 'update', target: 'x_remove',
      series: [{ name: 'B', values: [3, 4] }],
    }, removed.ast, removed.parser);
    expect(xValBlocks(chartXml(removed.parser))).toEqual([`<c:xVal>${numLitXml([5, 15])}</c:xVal>`]);
    const removedData = nodeChartData(removed.ast, 'x_remove') as {
      categories: string[];
      series: Array<{ name: string; xValues?: number[] }>;
    };
    expect(removedData.categories).toEqual(['5', '15']);
    expect(removedData.series.map((s) => [s.name, s.xValues])).toEqual([['B', [5, 15]]]);
    expect(renderChartText(chartNode(removed.ast, 'x_remove'))).toBe(
      '[Chart scatter] Latency vs load | B: (5, 3),(15, 4)'
    );
    const removedReopened = new DocxParser(removed.parser.rezip());
    const removedAst = await removedReopened.getAST();
    const removedNode = removedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(removedNode)).toBe('[Chart scatter] Latency vs load | B: (5, 3),(15, 4)');

    const reordered = await insertScatter('x_reorder', divergent);
    executeChart({
      action: 'chart', operation: 'update', target: 'x_reorder',
      series: [
        { name: 'B', values: [3, 4] },
        { name: 'A', values: [1, 2] },
      ],
    }, reordered.ast, reordered.parser);
    expect(xValBlocks(chartXml(reordered.parser))).toEqual([
      `<c:xVal>${numLitXml([5, 15])}</c:xVal>`,
      `<c:xVal>${numLitXml([0, 10])}</c:xVal>`,
    ]);
    const reorderedData = nodeChartData(reordered.ast, 'x_reorder') as { categories: string[] };
    expect(reorderedData.categories).toEqual(['5', '15']);
    expect(renderChartText(chartNode(reordered.ast, 'x_reorder'))).toBe(
      '[Chart scatter] Latency vs load | B: (5, 3),(15, 4) | A: (0, 1),(10, 2)'
    );
    const reorderedReopened = new DocxParser(reordered.parser.rezip());
    const reorderedAst = await reorderedReopened.getAST();
    const reorderedNode = reorderedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(reorderedNode)).toBe(
      '[Chart scatter] Latency vs load | B: (5, 3),(15, 4) | A: (0, 1),(10, 2)'
    );

    const prepended = await insertScatter('x_prepend', divergent);
    const before = chartXml(prepended.parser);
    const version = chartNode(prepended.ast, 'x_prepend').version;
    const sheetBefore = workbookSheet(prepended.parser);
    const metadataBefore = JSON.stringify(chartNode(prepended.ast, 'x_prepend').metadata);
    let err: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({
        action: 'chart', operation: 'update', target: 'x_prepend',
        series: [
          { name: 'C', values: [7, 8] },
          { name: 'A', values: [1, 2] },
          { name: 'B', values: [3, 4] },
        ],
      }, prepended.ast, prepended.parser);
    } catch (caught) {
      err = caught as Error & { code?: string; remedy?: string };
    }
    expect(err?.code).toBe('E_VALIDATION');
    expect(err?.message).toContain('different x lists');
    expect(err?.remedy).toContain('pass xValues');
    expect(chartXml(prepended.parser)).toBe(before);
    expect(chartNode(prepended.ast, 'x_prepend').version).toBe(version);
    expect(JSON.stringify(chartNode(prepended.ast, 'x_prepend').metadata)).toBe(metadataBefore);
    expect(workbookSheet(prepended.parser)).toBe(sheetBefore);

    executeChart({
      action: 'chart', operation: 'update', target: 'x_prepend',
      series: [
        { name: 'C', values: [7, 8], xValues: [7, 17] },
        { name: 'A', values: [1, 2] },
        { name: 'B', values: [3, 4] },
      ],
    }, prepended.ast, prepended.parser);
    expect(xValBlocks(chartXml(prepended.parser))).toEqual([
      `<c:xVal>${numLitXml([7, 17])}</c:xVal>`,
      `<c:xVal>${numLitXml([0, 10])}</c:xVal>`,
      `<c:xVal>${numLitXml([5, 15])}</c:xVal>`,
    ]);
    const prependedReopened = new DocxParser(prepended.parser.rezip());
    const prependedAst = await prependedReopened.getAST();
    const prependedNode = prependedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(prependedNode)).toBe(
      '[Chart scatter] Latency vs load | C: (7, 7),(17, 8) | A: (0, 1),(10, 2) | B: (5, 3),(15, 4)'
    );
  });

  it('refuses a repeated requested series name without categories before any write', async () => {
    const base = await insertScatter('dup_request', {
      categories: undefined,
      series: [
        { name: 'A', values: [1, 2], xValues: [0, 10] },
        { name: 'B', values: [3, 4], xValues: [5, 15] },
      ],
    });
    const node = chartNode(base.ast, 'dup_request');
    const before = chartXml(base.parser);
    const zipBefore = base.parser.rezip();
    const version = node.version;
    const metadataBefore = JSON.stringify(node.metadata);
    const sheetBefore = workbookSheet(base.parser);
    let err: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({
        action: 'chart', operation: 'update', target: 'dup_request',
        series: [
          { name: 'B', values: [7, 8] },
          { name: 'A', values: [1, 2] },
          { name: 'B', values: [3, 4] },
        ],
      }, base.ast, base.parser);
    } catch (caught) {
      err = caught as Error & { code?: string; remedy?: string };
    }
    expect(err?.code).toBe('E_VALIDATION');
    expect(err?.message).toContain('repeats series name "B"');
    expect(err?.remedy).toContain('xValues');
    expect(chartXml(base.parser)).toBe(before);
    expect(base.parser.rezip().equals(zipBefore)).toBe(true);
    expect(node.version).toBe(version);
    expect(JSON.stringify(node.metadata)).toBe(metadataBefore);
    expect(workbookSheet(base.parser)).toBe(sheetBefore);

    executeChart({
      action: 'chart', operation: 'update', target: 'dup_request',
      series: [
        { name: 'B', values: [7, 8], xValues: [5, 15] },
        { name: 'A', values: [1, 2], xValues: [0, 10] },
        { name: 'B', values: [3, 4], xValues: [7, 17] },
      ],
    }, base.ast, base.parser);
    expect(renderChartText(chartNode(base.ast, 'dup_request'))).toBe(
      '[Chart scatter] Latency vs load | B: (5, 7),(15, 8) | A: (0, 1),(10, 2) | B: (7, 3),(17, 4)'
    );
  });

  it('refuses duplicate stored series names without xValues and keeps the part unchanged', async () => {
    const base = await insertScatter('dup_stored');
    base.parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      mapSeries(base.xml, 1, (block) => block.replace('<c:tx><c:v>P99</c:v></c:tx>', '<c:tx><c:v>P50</c:v></c:tx>'))
    );
    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    const before = chartXml(reopened);
    const zipBefore = reopened.rezip();
    const version = node.version;
    const metadataBefore = JSON.stringify(node.metadata);
    const sheetBefore = workbookSheet(reopened);
    let err: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({
        action: 'chart', operation: 'update', target: node.id,
        series: [{ name: 'P50', values: [5, 15, 25] }],
      }, ast, reopened);
    } catch (caught) {
      err = caught as Error & { code?: string; remedy?: string };
    }
    expect(err?.code).toBe('E_VALIDATION');
    expect(err?.message).toContain('has 2 series named "P50"');
    expect(err?.remedy).toContain('xValues');
    expect(chartXml(reopened)).toBe(before);
    expect(reopened.rezip().equals(zipBefore)).toBe(true);
    expect(node.version).toBe(version);
    expect(JSON.stringify(node.metadata)).toBe(metadataBefore);
    expect(workbookSheet(reopened)).toBe(sheetBefore);

    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'P50', values: [5, 15, 25], xValues: [0, 10, 20] }],
    }, ast, reopened);
    expect(renderChartText(node)).toBe('[Chart scatter] Latency vs load | P50: (0, 5),(10, 15),(20, 25)');
  });

  it('carries shared x through a rename and refuses a divergent rename without xValues', async () => {
    const shared = await insertScatter('rename_shared');
    executeChart({
      action: 'chart', operation: 'update', target: 'rename_shared',
      series: [
        { name: 'P50 renamed', values: [5, 15, 25] },
        { name: 'P99 renamed', values: [12, 34, 58] },
      ],
    }, shared.ast, shared.parser);
    expect(xValBlocks(chartXml(shared.parser))).toEqual([
      `<c:xVal>${numLitXml([0, 10, 20])}</c:xVal>`,
      `<c:xVal>${numLitXml([0, 10, 20])}</c:xVal>`,
    ]);
    expect(renderChartText(chartNode(shared.ast, 'rename_shared'))).toBe(
      '[Chart scatter] Latency vs load | P50 renamed: (0, 5),(10, 15),(20, 25) | P99 renamed: (0, 12),(10, 34),(20, 58)'
    );
    const sharedReopened = new DocxParser(shared.parser.rezip());
    const sharedAst = await sharedReopened.getAST();
    const sharedNode = sharedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(sharedNode)).toBe(
      '[Chart scatter] Latency vs load | P50 renamed: (0, 5),(10, 15),(20, 25) | P99 renamed: (0, 12),(10, 34),(20, 58)'
    );

    const divergent = await insertScatter('rename_divergent', {
      categories: undefined,
      series: [
        { name: 'A', values: [1, 2], xValues: [0, 10] },
        { name: 'B', values: [3, 4], xValues: [5, 15] },
      ],
    });
    const node = chartNode(divergent.ast, 'rename_divergent');
    const before = chartXml(divergent.parser);
    const version = node.version;
    let err: (Error & { code?: string; remedy?: string }) | undefined;
    try {
      executeChart({
        action: 'chart', operation: 'update', target: 'rename_divergent',
        series: [{ name: 'A renamed', values: [1, 2] }],
      }, divergent.ast, divergent.parser);
    } catch (caught) {
      err = caught as Error & { code?: string; remedy?: string };
    }
    expect(err?.code).toBe('E_VALIDATION');
    expect(err?.message).toContain('different x lists');
    expect(err?.remedy).toContain('xValues');
    expect(chartXml(divergent.parser)).toBe(before);
    expect(node.version).toBe(version);

    executeChart({
      action: 'chart', operation: 'update', target: 'rename_divergent',
      series: [
        { name: 'A renamed', values: [1, 2], xValues: [0, 10] },
        { name: 'B', values: [3, 4] },
      ],
    }, divergent.ast, divergent.parser);
    expect(renderChartText(chartNode(divergent.ast, 'rename_divergent'))).toBe(
      '[Chart scatter] Latency vs load | A renamed: (0, 1),(10, 2) | B: (5, 3),(15, 4)'
    );
    const divergentReopened = new DocxParser(divergent.parser.rezip());
    const divergentAst = await divergentReopened.getAST();
    const divergentNode = divergentAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(renderChartText(divergentNode)).toBe(
      '[Chart scatter] Latency vs load | A renamed: (0, 1),(10, 2) | B: (5, 3),(15, 4)'
    );
  });

  it('keeps a titleless part titleless through structural regeneration', async () => {
    const count = await insertScatter('titleless_count');
    count.parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      count.xml.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const countReopened = new DocxParser(count.parser.rezip());
    const countAst = await countReopened.getAST();
    const countNode = countAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    executeChart({
      action: 'chart', operation: 'update', target: countNode.id,
      series: [{ name: 'P50', values: [6, 16, 26] }],
    }, countAst, countReopened);
    expect(chartXml(countReopened)).not.toContain('<c:title>');
    expect(Object.hasOwn(countNode.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(countNode.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(countNode)).toBe('[Chart scatter] | P50: (0, 6),(10, 16),(20, 26)');
    const countReopen = new DocxParser(countReopened.rezip());
    const countAst2 = await countReopen.getAST();
    const countNode2 = countAst2.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(Object.hasOwn(countNode2.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(countNode2.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(countNode2)).toBe('[Chart scatter] | P50: (0, 6),(10, 16),(20, 26)');
    expect(convertAstToHtml(countAst2, 'titleless_count')).not.toContain('Chart Analysis');

    const column = DocxParser.createNewDocument('blank');
    const columnAst = await column.getAST();
    executeChart({
      action: 'chart', chartType: 'column', id: 'titleless_switch', title: 'Temp',
      categories: ['0', '10', '20'], series: [{ name: 'S', values: [1, 2, 3] }],
    }, columnAst, column);
    column.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      chartXml(column).replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const switchReopened = new DocxParser(column.rezip());
    const switchAst = await switchReopened.getAST();
    const switchNode = switchAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    executeChart({
      action: 'chart', operation: 'update', target: switchNode.id, chartType: 'scatter',
    }, switchAst, switchReopened);
    expect(chartXml(switchReopened)).not.toContain('<c:title>');
    expect(Object.hasOwn(switchNode.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(switchNode.metadata.chartData!, 'title')).toBe(false);
    const switchData = switchNode.metadata.chartData as { series: Array<{ xValues?: number[] }> };
    expect(switchData.series[0].xValues).toEqual([0, 10, 20]);
    expect(renderChartText(switchNode)).toBe('[Chart scatter] | S: (0, 1),(10, 2),(20, 3)');
    const switchReopen = new DocxParser(switchReopened.rezip());
    const switchAst2 = await switchReopen.getAST();
    const switchNode2 = switchAst2.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(Object.hasOwn(switchNode2.metadata, 'title')).toBe(false);
    expect(renderChartText(switchNode2)).toBe('[Chart scatter] | S: (0, 1),(10, 2),(20, 3)');
    expect(convertAstToHtml(switchAst2, 'titleless_switch')).not.toContain('Chart Analysis');

    const generator = new Chart2DManipulator({ sections: [], relationships: {} } as never);
    const explicitEmpty = generator.generate2DChartXml({
      chartType: 'column', title: '', categories: ['Q1'], series: [{ name: 'S', values: [1] }],
    });
    expect(explicitEmpty).not.toContain('<c:title>');
    const defaulted = generator.generate2DChartXml({
      chartType: 'column', categories: ['Q1'], series: [{ name: 'S', values: [1] }],
    });
    expect(defaulted).toContain('<a:t>Chart Analysis</a:t>');

    const threeD = DocxParser.createNewDocument('blank');
    const threeDAst = await threeD.getAST();
    executeChart({
      action: 'chart', chartType: 'bar3D', id: 'titleless_3d', title: 'Temp 3D',
      categories: ['Q1', 'Q2'], series: [{ name: 'S', values: [1, 2] }, { name: 'T', values: [3, 4] }],
    }, threeDAst, threeD);
    threeD.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      chartXml(threeD).replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const threeDReopened = new DocxParser(threeD.rezip());
    const threeDAst2 = await threeDReopened.getAST();
    const threeDNode = threeDAst2.sections[0].nodes.find((n) => n.type === 'chart3d')!;
    executeChart({
      action: 'chart', operation: 'update', target: threeDNode.id,
      series: [{ name: 'S', values: [4, 5] }],
    }, threeDAst2, threeDReopened);
    expect(chartXml(threeDReopened)).not.toContain('<c:title>');
    expect(Object.hasOwn(threeDNode.metadata, 'title')).toBe(false);
    expect(renderChartText(threeDNode)).not.toContain('3D Chart Analysis');
    const threeDReopen = new DocxParser(threeDReopened.rezip());
    const threeDAst3 = await threeDReopen.getAST();
    const threeDNode3 = threeDAst3.sections[0].nodes.find((n) => n.type === 'chart3d')!;
    expect(chartXml(threeDReopen)).not.toContain('<c:title>');
    expect(Object.hasOwn(threeDNode3.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(threeDNode3.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(threeDNode3)).not.toContain('3D Chart Analysis');
  });

  it('removes the chart title for an explicit empty title in place', async () => {
    const titled = await insertScatter('clear_title');
    expect(chartXml(titled.parser)).toContain('<c:title>');
    executeChart({
      action: 'chart', operation: 'update', target: 'clear_title', title: '',
    }, titled.ast, titled.parser);
    const node = chartNode(titled.ast, 'clear_title');
    expect(chartXml(titled.parser)).not.toContain('<c:title');
    expect(Object.hasOwn(node.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(node.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(node)).toBe(
      '[Chart scatter] | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58)'
    );
    expect(convertAstToHtml(titled.ast, 'clear_title')).not.toContain('Latency vs load');
    const titledReopened = new DocxParser(titled.parser.rezip());
    const titledAst = await titledReopened.getAST();
    const titledNode = titledAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chartXml(titledReopened)).not.toContain('<c:title');
    expect(Object.hasOwn(titledNode.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(titledNode.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(titledNode)).toBe(
      '[Chart scatter] | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58)'
    );
    expect(convertAstToHtml(titledAst, 'clear_title')).not.toContain('Latency vs load');

    const titleless = await insertScatter('clear_titleless');
    titleless.parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      titleless.xml.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
    );
    const titlelessReopened = new DocxParser(titleless.parser.rezip());
    const titlelessAst = await titlelessReopened.getAST();
    const titlelessNode = titlelessAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const before = chartXml(titlelessReopened);
    executeChart({
      action: 'chart', operation: 'update', target: titlelessNode.id, title: '',
    }, titlelessAst, titlelessReopened);
    expect(chartXml(titlelessReopened)).toBe(before);
    expect(chartXml(titlelessReopened)).not.toContain('<c:title');
    expect(Object.hasOwn(titlelessNode.metadata, 'title')).toBe(false);
    expect(renderChartText(titlelessNode)).toBe(
      '[Chart scatter] | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58)'
    );
    expect(convertAstToHtml(titlelessAst, 'clear_titleless')).not.toContain('Chart Analysis');
    const titlelessAgain = new DocxParser(titlelessReopened.rezip());
    const titlelessAstAgain = await titlelessAgain.getAST();
    const titlelessNodeAgain = titlelessAstAgain.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chartXml(titlelessAgain)).not.toContain('<c:title');
    expect(Object.hasOwn(titlelessNodeAgain.metadata, 'title')).toBe(false);
    expect(renderChartText(titlelessNodeAgain)).not.toContain('Chart Analysis');
  });

  it('never reads an axis title as the chart title when an explicit empty title clears the chart title', async () => {
    const inserted = await insertScatter('clear_axis_title');
    const axisTitle = '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Y AXIS KEEP</a:t></a:r></a:p></c:rich></c:tx></c:title>';
    expect(inserted.xml).toContain('<c:valAx>');
    const withAxis = inserted.xml.replace(/(<c:valAx>[\s\S]*?)<c:numFmt/, `$1${axisTitle}<c:numFmt`);
    expect(withAxis).toContain(axisTitle);
    inserted.parser.getZipAdapter().writeText('word/charts/chart1.xml', withAxis);

    const reopened = new DocxParser(inserted.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chartLevelXml(chartXml(reopened))).toContain('Latency vs load');
    expect(node.metadata.chartData!.title).toBe('Latency vs load');

    executeChart({
      action: 'chart', operation: 'update', target: node.id, title: '',
    }, ast, reopened);
    const after = chartXml(reopened);
    expect(chartLevelXml(after)).not.toContain('<c:title');
    expect(after).toContain(axisTitle);
    expect(summarize2DChartXml(after).title).toBeUndefined();
    expect(Object.hasOwn(node.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(node.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(node)).toBe(
      '[Chart scatter] | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58)'
    );
    expect(renderChartText(node)).not.toContain('Y AXIS KEEP');
    expect(convertAstToHtml(ast, 'clear_axis_title')).not.toContain('Y AXIS KEEP');
    expect(convertAstToHtml(ast, 'clear_axis_title')).not.toContain('Latency vs load');
    expect(summarize2DChartXml(toDefaultNamespace(after)).title).toBeUndefined();
    expect(summarize2DChartXml(toXAlias(after)).title).toBeUndefined();
    const numericPrefix = after
      .replace(`xmlns:c="${CHART_NS_URI}"`, `xmlns:ns0="${CHART_NS_URI}"`)
      .replace(/<(\/?)c:/g, '<$1ns0:');
    expect(summarize2DChartXml(numericPrefix).title).toBeUndefined();

    // A nested <c:title> before the plot area is not a legal chart title
    // source (CT_Chart puts it directly under <c:chart>); it must not be read.
    const nestedTitle = after.replace(
      '<c:plotArea',
      '<c:extLst><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>VENDOR TITLE</a:t></a:r></a:p></c:rich></c:tx></c:title></c:extLst><c:plotArea'
    );
    expect(nestedTitle).toContain('VENDOR TITLE');
    expect(summarize2DChartXml(nestedTitle).title).toBeUndefined();

    // Reopen each namespace form through the parser hydration path, not just
    // the summary: an axis title must never become chartData.title.
    for (const [label, nsXml] of [
      ['default', toDefaultNamespace(after)],
      ['alias', toXAlias(after)],
      ['numeric', numericPrefix],
      ['nested', nestedTitle],
    ] as const) {
      const nsParser = new DocxParser(reopened.rezip());
      nsParser.getZipAdapter().writeText('word/charts/chart1.xml', nsXml);
      const nsAst = await nsParser.getAST();
      const nsNode = nsAst.sections[0].nodes.find((n) => n.type === 'chart')!;
      expect(Object.hasOwn(nsNode.metadata.chartData!, 'title')).toBe(false);
      expect(renderChartText(nsNode)).not.toContain('Y AXIS KEEP');
      expect(renderChartText(nsNode)).not.toContain('VENDOR TITLE');
      const nsHtml = convertAstToHtml(nsAst, `clear_axis_title_${label}`);
      expect(nsHtml).not.toContain('Y AXIS KEEP');
      expect(nsHtml).not.toContain('VENDOR TITLE');
    }

    // Writer and reader share the direct-child bound: an explicit title on a
    // part whose only chart-level <c:title> is a hostile nested one must write
    // a real direct title, not silently patch the nested element.
    const nestedParser = new DocxParser(reopened.rezip());
    nestedParser.getZipAdapter().writeText('word/charts/chart1.xml', nestedTitle);
    const nestedAst = await nestedParser.getAST();
    const nestedNode = nestedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    executeChart({
      action: 'chart', operation: 'update', target: nestedNode.id, title: 'NEW T',
    }, nestedAst, nestedParser);
    const nestedAfter = chartXml(nestedParser);
    // Strip the hostile extLst: what remains in the chart-level region must be
    // a real direct title, not the nested vendor element.
    const directRegion = chartLevelXml(nestedAfter).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
    expect(directRegion).toContain('<c:title');
    expect(directRegion).toContain('NEW T');
    expect(summarize2DChartXml(nestedAfter).title).toBe('NEW T');
    expect(renderChartText(nestedNode)).toContain('NEW T');
    expect(renderChartText(nestedNode)).not.toContain('VENDOR TITLE');
    expect(convertAstToHtml(nestedAst, 'nested_title_write')).not.toContain('VENDOR TITLE');

    const nestedAgain = new DocxParser(nestedParser.rezip());
    const nestedAstAgain = await nestedAgain.getAST();
    const nestedNodeAgain = nestedAstAgain.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(nestedNodeAgain.metadata.chartData!.title).toBe('NEW T');
    expect(summarize2DChartXml(chartXml(nestedAgain)).title).toBe('NEW T');
    expect(renderChartText(nestedNodeAgain)).toContain('NEW T');
    expect(renderChartText(nestedNodeAgain)).not.toContain('VENDOR TITLE');

    // A hostile duplicate direct title must not resurface after a clear.
    const duplicateParser = new DocxParser(inserted.parser.rezip());
    const duplicateXml = inserted.xml.replace(/(<c:title>[\s\S]*?<\/c:title>)/, '$1$1');
    expect(duplicateXml.match(/<c:title>/g)!.length).toBe(2);
    duplicateParser.getZipAdapter().writeText('word/charts/chart1.xml', duplicateXml);
    const duplicateAst = await duplicateParser.getAST();
    const duplicateNode = duplicateAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    executeChart({
      action: 'chart', operation: 'update', target: duplicateNode.id, title: '',
    }, duplicateAst, duplicateParser);
    const duplicateAfter = chartXml(duplicateParser);
    expect(chartLevelXml(duplicateAfter)).not.toContain('<c:title');
    expect(summarize2DChartXml(duplicateAfter).title).toBeUndefined();
    expect(renderChartText(duplicateNode)).not.toContain('Latency vs load');

    // A hostile nested autoTitleDeleted marker before plotArea must not become
    // the insertion anchor; the direct title is written and the marker is kept.
    const markerParser = new DocxParser(reopened.rezip());
    const markerXml = after.replace(
      '<c:plotArea',
      '<c:extLst><c:autoTitleDeleted val="9"/></c:extLst><c:plotArea'
    );
    markerParser.getZipAdapter().writeText('word/charts/chart1.xml', markerXml);
    const markerAst = await markerParser.getAST();
    const markerNode = markerAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    executeChart({
      action: 'chart', operation: 'update', target: markerNode.id, title: 'MARKER T',
    }, markerAst, markerParser);
    const markerAfter = chartXml(markerParser);
    const markerDirect = chartLevelXml(markerAfter).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
    expect(markerDirect).toContain('<c:title');
    expect(markerDirect).toContain('MARKER T');
    expect(markerAfter).toContain('<c:autoTitleDeleted val="9"/>');
    expect(summarize2DChartXml(markerAfter).title).toBe('MARKER T');
    const markerAgain = new DocxParser(markerParser.rezip());
    const markerAstAgain = await markerAgain.getAST();
    const markerNodeAgain = markerAstAgain.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(markerNodeAgain.metadata.chartData!.title).toBe('MARKER T');
    expect(renderChartText(markerNodeAgain)).toContain('MARKER T');

    const reopenedAgain = new DocxParser(reopened.rezip());
    const astAgain = await reopenedAgain.getAST();
    const nodeAgain = astAgain.sections[0].nodes.find((n) => n.type === 'chart')!;
    const afterReopen = chartXml(reopenedAgain);
    expect(chartLevelXml(afterReopen)).not.toContain('<c:title');
    expect(afterReopen).toContain(axisTitle);
    expect(summarize2DChartXml(afterReopen).title).toBeUndefined();
    expect(Object.hasOwn(nodeAgain.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(nodeAgain.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(nodeAgain)).toBe(
      '[Chart scatter] | P50: (0, 5),(10, 15),(20, 25) | P99: (0, 12),(10, 34),(20, 58)'
    );
    expect(renderChartText(nodeAgain)).not.toContain('Y AXIS KEEP');
    expect(convertAstToHtml(astAgain, 'clear_axis_title')).not.toContain('Y AXIS KEEP');
    expect(convertAstToHtml(astAgain, 'clear_axis_title')).not.toContain('Latency vs load');

    executeChart({
      action: 'chart', operation: 'update', target: nodeAgain.id,
      series: [
        { name: 'P50', values: [6, 16, 26] },
        { name: 'P99', values: [13, 35, 59] },
      ],
    }, astAgain, reopenedAgain);
    const inPlace = chartXml(reopenedAgain);
    expect(chartLevelXml(inPlace)).not.toContain('<c:title');
    expect(inPlace).toContain(axisTitle);
    expect(summarize2DChartXml(inPlace).title).toBeUndefined();
    expect(Object.hasOwn(nodeAgain.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(nodeAgain)).toBe(
      '[Chart scatter] | P50: (0, 6),(10, 16),(20, 26) | P99: (0, 13),(10, 35),(20, 59)'
    );

    executeChart({
      action: 'chart', operation: 'update', target: nodeAgain.id,
      series: [{ name: 'P50', values: [6, 16, 26] }],
    }, astAgain, reopenedAgain);
    const structural = chartXml(reopenedAgain);
    // A series-count change rebuilds the part (historical regeneration path),
    // so the unmodeled axis title XML is not carried over; what must never
    // happen is its text reappearing as a chart-level title.
    expect(chartLevelXml(structural)).not.toContain('<c:title');
    expect(structural).not.toContain('Y AXIS KEEP');
    expect(summarize2DChartXml(structural).title).toBeUndefined();
    expect(Object.hasOwn(nodeAgain.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(nodeAgain.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(nodeAgain)).toBe('[Chart scatter] | P50: (0, 6),(10, 16),(20, 26)');
    expect(renderChartText(nodeAgain)).not.toContain('Y AXIS KEEP');
    expect(convertAstToHtml(astAgain, 'clear_axis_title')).not.toContain('Y AXIS KEEP');

    const structuralReopened = new DocxParser(reopenedAgain.rezip());
    const structuralAst = await structuralReopened.getAST();
    const structuralNode = structuralAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const structuralXml = chartXml(structuralReopened);
    expect(chartLevelXml(structuralXml)).not.toContain('<c:title');
    expect(structuralXml).not.toContain('Y AXIS KEEP');
    expect(summarize2DChartXml(structuralXml).title).toBeUndefined();
    expect(Object.hasOwn(structuralNode.metadata, 'title')).toBe(false);
    expect(Object.hasOwn(structuralNode.metadata.chartData!, 'title')).toBe(false);
    expect(renderChartText(structuralNode)).toBe('[Chart scatter] | P50: (0, 6),(10, 16),(20, 26)');
    expect(renderChartText(structuralNode)).not.toContain('Y AXIS KEEP');
    expect(convertAstToHtml(structuralAst, 'clear_axis_title')).not.toContain('Y AXIS KEEP');
  });

  it('retitles a sparse scatter part without touching its x or y data', async () => {
    const { parser } = await insertScatter('sparse_retitle');
    const sparseXml = chartXml(parser).replace(/(<c:xVal>[\s\S]*?<\/c:xVal>)/, (block) =>
      block.replace('idx="1"', 'idx="5"')
    );
    parser.getZipAdapter().writeText('word/charts/chart1.xml', sparseXml);
    const reopened = new DocxParser(parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    const before = chartXml(reopened);
    const sheetBefore = workbookSheet(reopened);
    const version = node.version;
    executeChart({
      action: 'chart', operation: 'update', target: node.id, title: 'Sparse Retitled',
    }, ast, reopened);
    const after = chartXml(reopened);
    expect(after).toContain('<a:t>Sparse Retitled</a:t>');
    expect(xValBlocks(after)).toEqual(xValBlocks(before));
    expect(yValBlocks(after)).toEqual(yValBlocks(before));
    expect(workbookSheet(reopened)).toBe(sheetBefore);
    expect(node.version).toBe(version + 1);
    expect((node.metadata.chartData as { categories: string[] }).categories).toEqual(['0', '20']);
    expect(renderChartText(node)).toContain('P50: (0, 5),(20, 25); 1 y value without x');
  });

  it('compacts sparse x metadata and reports unpaired x and y counts', async () => {
    const sparsePayload = {
      title: 'Sparse counts',
      categories: undefined,
      series: [
        { name: 'SX', values: [10, 20, 30], xValues: [1, 2, 3] },
        { name: 'SY', values: [7, 8, 9], xValues: [1, 2, 3] },
      ],
    };

    const sparseX = await insertScatter('sparse_x_meta', sparsePayload);
    const sparseXPart = mapSeries(sparseX.xml, 0, (block) =>
      block.replace(
        /<c:xVal><c:numLit>[\s\S]*?<\/c:numLit><\/c:xVal>/,
        `<c:xVal>${sparseNumLitXml(3, [[0, 1], [2, 3]])}</c:xVal>`
      )
    );
    sparseX.parser.getZipAdapter().writeText('word/charts/chart1.xml', sparseXPart);
    const xReopened = new DocxParser(sparseX.parser.rezip());
    const xAst = await xReopened.getAST();
    const xNode = xAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const xData = xNode.metadata.chartData as {
      categories: string[];
      series: Array<{ name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number }>;
    };
    expect(xData.categories).toEqual(['1', '3']);
    expect(xData.series[0]).toMatchObject({ values: [10, 30], xValues: [1, 3], unpairedY: 1 });
    expect(xData.series[0].unpairedX).toBeUndefined();
    expect(renderChartText(xNode)).toBe(
      '[Chart scatter] Sparse counts | SX: (1, 10),(3, 30); 1 y value without x | SY: (1, 7),(2, 8),(3, 9)'
    );
    const xHtml = convertAstToHtml(xAst, 'sparse_x_meta');
    expect(xHtml).toContain('<title>SX: (3, 30)</title>');
    expect(xHtml).not.toContain('<title>SX: (2, 20)</title>');
    expect(xHtml).toContain('1 point omitted: no x value');
    const xSummary = summarize2DChartXml(chartXml(xReopened));
    expect(xSummary.categories).toEqual(['1', '3']);
    expect(xSummary.series[0]).toMatchObject({
      values: [10, 30], xValues: [1, 3], unpairedY: 1,
    });

    const sparseY = await insertScatter('sparse_y_meta', sparsePayload);
    const sparseYPart = mapSeries(sparseY.xml, 0, (block) =>
      block.replace(
        /<c:yVal><c:numLit>[\s\S]*?<\/c:numLit><\/c:yVal>/,
        `<c:yVal>${sparseNumLitXml(3, [[0, 10], [2, 30]])}</c:yVal>`
      )
    );
    sparseY.parser.getZipAdapter().writeText('word/charts/chart1.xml', sparseYPart);
    const yReopened = new DocxParser(sparseY.parser.rezip());
    const yAst = await yReopened.getAST();
    const yNode = yAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const yData = yNode.metadata.chartData as {
      categories: string[];
      series: Array<{ name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number }>;
    };
    expect(yData.categories).toEqual(['1', '3']);
    expect(yData.series[0]).toMatchObject({ values: [10, 30], xValues: [1, 3], unpairedX: 1 });
    expect(yData.series[0].unpairedY).toBeUndefined();
    expect(renderChartText(yNode)).toBe(
      '[Chart scatter] Sparse counts | SX: (1, 10),(3, 30); 1 x value without y | SY: (1, 7),(2, 8),(3, 9)'
    );
    const yHtml = convertAstToHtml(yAst, 'sparse_y_meta');
    expect(yHtml).toContain('<title>SX: (3, 30)</title>');
    expect(yHtml).not.toContain('<title>SX: (2, 20)</title>');
    expect(yHtml).toContain('1 point omitted: no y value');
    const ySummary = summarize2DChartXml(chartXml(yReopened));
    expect(ySummary.categories).toEqual(['1', '3']);
    expect(ySummary.series[0]).toMatchObject({
      values: [10, 30], xValues: [1, 3], unpairedX: 1,
    });
  });

  it('exposes x-only scatter series and reports every unpaired x point', async () => {
    const base = await insertScatter('x_only', {
      categories: ['1', '2', '3'],
      series: [{ name: 'S', values: [10, 20, 30] }],
    });
    const part = mapSeries(base.xml, 0, (block) => block.replace(/<c:yVal>[\s\S]*?<\/c:yVal>/, ''));
    expect(yValBlocks(part)).toHaveLength(0);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', part);
    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    const data = node.metadata.chartData as {
      categories: string[];
      series: Array<{ name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number }>;
    };
    expect(data.categories).toEqual(['1', '2', '3']);
    expect(data.series[0]).toMatchObject({ name: 'S', values: [], xValues: [1, 2, 3], unpairedX: 3 });
    expect(data.series[0].unpairedY).toBeUndefined();
    expect(renderChartText(node)).toBe('[Chart scatter] Latency vs load | S: 3 x values without y');
    const html = convertAstToHtml(ast, 'x_only');
    expect(html).toContain('3 points omitted: no y value');
    expect(html).not.toContain('<title>S: (');
    const summary = summarize2DChartXml(chartXml(reopened));
    expect(summary.categories).toEqual(['1', '2', '3']);
    expect(summary.series[0]).toMatchObject({ values: [], xValues: [1, 2, 3], unpairedX: 3 });
    expect(summary.series[0].unpairedY).toBeUndefined();

    executeChart({
      action: 'chart', operation: 'update', target: node.id, title: 'X-only retitle',
    }, ast, reopened);
    const afterUpdate = node.metadata.chartData as {
      series: Array<{ values: number[]; xValues?: number[]; unpairedX?: number }>;
    };
    expect(afterUpdate.series[0]).toMatchObject({ values: [], xValues: [1, 2, 3], unpairedX: 3 });
    expect(renderChartText(node)).toBe('[Chart scatter] X-only retitle | S: 3 x values without y');
  });
});
