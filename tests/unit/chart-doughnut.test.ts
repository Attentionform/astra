import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import * as crypto from 'crypto';
import {
  Chart2DManipulator,
  Chart3DManipulator,
  DocxParser,
  DocumentAST,
  OperationSchema,
  SUPPORTED_CHART_2D_AND_3D_TYPES,
  ToonGenerator,
  chartHoleSizeSchemaRemedy,
  convertAstToHtml,
  detectBar2DChartTypeFromXml,
  detectChartKindFromXml,
  executeChart,
  renderChartText,
  summarize2DChartXml,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

/**
 * Doughnut pins: pie-family writer with c:holeSize, reader/detection, insert,
 * pie-equivalent update routing (a plain 'pie' request must never flatten an
 * existing doughnut), foreign styling preservation, TOON, HTML and validity.
 */

const DEMO = {
  title: 'Revenue by quarter',
  categories: ['Q1', 'Q2', 'Q3'],
  series: [{ name: 'ARR', values: [10, 20, 30], color: '0066CC' }],
};

const CHART_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** Rewrites the chart part to a default chart namespace with c: bound to a vendor. */
function toDefaultNamespace(xml: string): string {
  return xml
    .replace(`xmlns:c="${CHART_NS_URI}"`, `xmlns="${CHART_NS_URI}" xmlns:c="urn:example:vendor"`)
    .replace(/<(\/?)c:/g, '<$1');
}

/** Rewrites every chart-namespace tag to an x: alias. */
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

function chartNode(ast: DocumentAST, id: string) {
  return ast.sections[0].nodes.find((n) => n.id === id)!;
}

async function insertChart(chartType: 'doughnut' | 'pie' | 'column', id = `chart_${chartType}`, opts: Record<string, unknown> = {}) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  executeChart({ action: 'chart', chartType, id, ...DEMO, ...opts }, ast, parser);
  const buf = parser.rezip();
  return { parser, ast, id, buf, xml: chartXmlOf(buf) };
}

function chartXml(parser: DocxParser): string {
  return parser.getZipAdapter().readText('word/charts/chart1.xml');
}

function updateSeries(parser: DocxParser, ast: DocumentAST, id: string, chartType?: 'doughnut' | 'pie', values = [99, 20, 30]) {
  executeChart({
    action: 'chart',
    operation: 'update',
    target: id,
    ...(chartType ? { chartType } : {}),
    series: [{ name: 'ARR', values }],
  }, ast, parser);
}

function expectSingleValueSpan(before: string, after: string): void {
  const oldSpan = '<c:v>10</c:v>';
  const at = before.indexOf(oldSpan);
  expect(at).toBeGreaterThan(0);
  expect(after).toBe(before.slice(0, at) + '<c:v>99</c:v>' + before.slice(at + oldSpan.length));
}

describe('doughnut charts', () => {
  it('writes pie-family geometry with a default 50% hole in schema order', () => {
    const xml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'doughnut', ...DEMO });

    expect(xml).toContain('<c:doughnutChart>');
    expect(xml).toContain('</c:doughnutChart>');
    expect(xml).toContain('<c:varyColors val="1"/>');
    expect(xml).toContain('<c:showPercent val="1"/>');
    expect(xml).toContain('<c:firstSliceAng val="0"/>');
    expect(xml).toContain('<c:holeSize val="50"/>');
    expect(xml).toContain('<c:dPt>');
    expect(xml).toContain('<c:numCache>');
    expect(xml).not.toContain('<c:pieChart>');
    expect(xml).not.toContain('<c:catAx>');
    expect(xml).not.toContain('<c:valAx>');

    const order = [
      '<c:varyColors val="1"/>',
      '<c:ser>',
      '<c:dLbls>',
      '<c:firstSliceAng val="0"/>',
      '<c:holeSize val="50"/>',
    ].map((tag) => xml.indexOf(tag));
    for (let i = 1; i < order.length; i++) {
      expect(order[i]).toBeGreaterThan(order[i - 1]);
    }
  });

  it('keeps plain pie hole-free', () => {
    const xml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'pie', ...DEMO });
    expect(xml).toContain('<c:pieChart>');
    expect(xml).not.toContain('doughnut');
    expect(xml).not.toContain('<c:holeSize');
  });

  it('keeps legacy c: chart writer bytes identical to the pre-doughnut baseline', () => {
    const baseline: Record<string, string> = {
      column: 'b770508c2048e21c7780f9c7b1ec4115ab26a8034cf4b1c4a56850cbabe4f3e1',
      bar: 'f462502f6d15c1405d6a80601966576c52f88c17a017730ab448402c54f82dee',
      line: '4a302064e21fca58f46e87df8fd906fb46b1a77b18fa816a10ec89e629ec04df',
      pie: 'a1c4ac842515523e67c10269744e5a3057292ea36b7cdcafa48db9ee7e3b315c',
      area: 'c12a069f6da02bba8b4ea910152ff720f733022e27e028e1e2dd8d9fceeb2eb3',
      column3D: '58d895b545b35413be65b35afc89ba76240176741c9969f43c56be1499398758',
      bar3D: '1b37db9feff92129dd467a23a22d7af70573e088a8c9f50fc8469dbc58ae3ba6',
      line3D: '68929996f6eea1a85772c1210a962a42480f868f2b9c5b4edc1c3caaa790e2e5',
      pie3D: 'b649bd696283e663967de6e52dc0a3e941eab9d1bbc8b860086548924111975c',
    };
    const demo = {
      title: 'T',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'ARR', values: [10, 20, 30], color: '0066CC' },
        { name: 'Costs', values: [5, 7, 9], color: '10B981' },
      ],
    };
    const hash = (xml: string): string => crypto.createHash('sha256').update(xml).digest('hex');
    for (const type of ['column', 'bar', 'line', 'pie', 'area'] as const) {
      expect(hash(new Chart2DManipulator({ sections: [], relationships: {} } as never)
        .generate2DChartXml({ chartType: type, ...demo }))).toBe(baseline[type]);
    }
    for (const type of ['column3D', 'bar3D', 'line3D', 'pie3D'] as const) {
      expect(hash(new Chart3DManipulator({ sections: [], relationships: {} } as never)
        .generate3DChartXml({ chartType: type, ...demo }))).toBe(baseline[type]);
    }
  });

  it('inserts a valid doughnut part with matching metadata', async () => {
    const { parser, ast, id, buf, xml } = await insertChart('doughnut');
    expect(xml).toContain('<c:doughnutChart>');
    expect(xml).toContain('<c:holeSize val="50"/>');
    const node = chartNode(ast, id);
    expect(node.type).toBe('chart');
    expect(node.metadata.chartType).toBe('doughnut');
    expect((node.metadata.chartData as { chartType?: string }).chartType).toBe('doughnut');
    const docx = await validateDocxBuffer(buf);
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
    const ooxml = await parser.validateOOXML();
    expect(ooxml.ok).toBe(true);
    expect(ooxml.errors).toEqual([]);
  });

  it('detects and summarizes the doughnut family, leaving pie detection intact', async () => {
    const doughnut = await insertChart('doughnut');
    expect(detectChartKindFromXml(doughnut.xml)).toEqual({ is3D: false, chartType: 'doughnut' });
    expect(summarize2DChartXml(doughnut.xml).chartType).toBe('doughnut');
    expect(detectBar2DChartTypeFromXml(doughnut.xml)).toBeUndefined();

    const pie = await insertChart('pie');
    expect(detectChartKindFromXml(pie.xml)).toEqual({ is3D: false, chartType: 'pie' });
    expect(summarize2DChartXml(pie.xml).chartType).toBe('pie');
  });

  it('loads an existing doughnut part as a doughnut chart node', async () => {
    const { buf } = await insertChart('doughnut', 'loaded_chart');
    const reopened = new DocxParser(buf);
    const ast = await reopened.getAST();
    const node = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('doughnut');
    expect((node.metadata.chartData as { chartType?: string }).chartType).toBe('doughnut');
  });

  it('TOON surfaces the full doughnut name for read-to-update round-trips', async () => {
    const { ast, id } = await insertChart('doughnut');
    const node = chartNode(ast, id);
    expect(renderChartText(node)).toContain('[Chart doughnut]');
    const row = new ToonGenerator().toon(ast, { fidelity: 'standard' }).rows.find((r) => r.id === id)!;
    expect(row.flags).toContain('[CH:doughnut]');
  });

  it('operation schema accepts doughnut in the chart and legacy insert shapes', () => {
    expect(SUPPORTED_CHART_2D_AND_3D_TYPES).toContain('doughnut');
    expect(OperationSchema.safeParse({
      action: 'chart',
      operation: 'insert',
      anchor: 'p_001',
      chartType: 'doughnut',
      categories: ['Q1'],
      series: [{ name: 'S', values: [1] }],
    }).success).toBe(true);
    expect(OperationSchema.safeParse({
      action: 'insert',
      type: 'chart',
      anchor: 'p_001',
      position: 'after',
      chartType: 'doughnut',
      categories: ['Q1'],
      series: [{ name: 'S', values: [1] }],
    }).success).toBe(true);
    expect(OperationSchema.safeParse({
      action: 'chart',
      operation: 'insert',
      anchor: 'p_001',
      chartType: 'bubble',
      categories: ['Q1'],
      series: [{ name: 'S', values: [1] }],
    }).success).toBe(false);
  });

  it('values-only update with no chartType merges in place and keeps the hole', async () => {
    const { parser, ast, id } = await insertChart('doughnut');
    const before = chartXml(parser);
    updateSeries(parser, ast, id);
    const after = chartXml(parser);
    expectSingleValueSpan(before, after);
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="50"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe('doughnut');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('chartType pie on an existing doughnut is a family selector, never a flatten', async () => {
    const { parser, ast, id } = await insertChart('doughnut');
    const before = chartXml(parser);
    updateSeries(parser, ast, id, 'pie');
    const after = chartXml(parser);
    expectSingleValueSpan(before, after);
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="50"/>');
    expect(after).not.toContain('<c:pieChart>');
    expect(chartNode(ast, id).metadata.chartType).toBe('doughnut');
    expect(nodeChartData(ast, id).chartType).toBe('doughnut');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('explicit chartType doughnut on an existing doughnut merges in place', async () => {
    const { parser, ast, id } = await insertChart('doughnut');
    const before = chartXml(parser);
    updateSeries(parser, ast, id, 'doughnut');
    const after = chartXml(parser);
    expectSingleValueSpan(before, after);
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="50"/>');
  });

  it('preserves authored doughnut styling on a data-only update', async () => {
    const { parser, ast, id } = await insertChart('doughnut');
    const foreign = chartXml(parser)
      .replace('<c:holeSize val="50"/>', '<c:holeSize val="75"/>')
      .replace('<c:legendPos val="r"/>', '<c:legendPos val="t"/>')
      .replace('<c:showVal val="0"/>', '<c:showVal val="1"/>')
      .replace('<c:showPercent val="1"/>', '<c:showPercent val="0"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', foreign);
    updateSeries(parser, ast, id);
    const after = chartXml(parser);
    expectSingleValueSpan(foreign, after);
    expect(after).toContain('<c:holeSize val="75"/>');
    expect(after).toContain('<c:legendPos val="t"/>');
    expect(after).toContain('<c:showVal val="1"/>');
    expect(after).toContain('<c:showPercent val="0"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe('doughnut');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('a series-count change regenerates in the detected doughnut family', async () => {
    const { parser, ast, id } = await insertChart('doughnut');
    executeChart({
      action: 'chart',
      operation: 'update',
      target: id,
      series: [
        { name: 'ARR', values: [99, 20, 30] },
        { name: 'Costs', values: [5, 7, 9] },
      ],
    }, ast, parser);
    const after = chartXml(parser);
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="50"/>');
    expect(after).not.toContain('<c:pieChart>');
    expect(chartNode(ast, id).metadata.chartType).toBe('doughnut');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('explicit pie to doughnut converts shape and writes the hole', async () => {
    const { parser, ast, id } = await insertChart('pie');
    executeChart({ action: 'chart', operation: 'update', target: id, chartType: 'doughnut' }, ast, parser);
    const after = chartXml(parser);
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="50"/>');
    expect(after).not.toContain('<c:pieChart>');
    expect(chartNode(ast, id).metadata.chartType).toBe('doughnut');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('HTML export draws annular doughnut sectors instead of a pie with a hole', async () => {
    const doughnut = await insertChart('doughnut');
    const doughnutHtml = convertAstToHtml(doughnut.ast, 'doughnut');
    expect(doughnutHtml).toContain('A 80 80 0');
    expect(doughnutHtml).toContain('A 40 40 0');
    expect(doughnutHtml).not.toContain('M 170 120 L');

    const pie = await insertChart('pie');
    const pieHtml = convertAstToHtml(pie.ast, 'pie');
    expect(pieHtml).toContain('M 170 120 L');
    expect(pieHtml).not.toContain('A 40 40 0');
  });

  it('accepts an authored holeSize on insert and update, with 1..90 schema bounds', async () => {
    const inserted = await insertChart('doughnut', 'hole75', { holeSize: 75 });
    expect(inserted.xml).toContain('<c:holeSize val="75"/>');
    expect((nodeChartData(inserted.ast, 'hole75') as { holeSize?: number }).holeSize).toBe(75);

    const updated = await insertChart('doughnut', 'hole_update');
    executeChart({
      action: 'chart', operation: 'update', target: 'hole_update', holeSize: 60,
    }, updated.ast, updated.parser);
    expect(chartXml(updated.parser)).toContain('<c:holeSize val="60"/>');

    for (const good of [1, 50, 90]) {
      expect(OperationSchema.safeParse({ action: 'chart', operation: 'insert', anchor: 'p_001', chartType: 'doughnut', holeSize: good }).success).toBe(true);
    }
    for (const bad of [0, 91, 100, 50.5]) {
      expect(OperationSchema.safeParse({ action: 'chart', operation: 'insert', anchor: 'p_001', chartType: 'doughnut', holeSize: bad }).success).toBe(false);
      expect(OperationSchema.safeParse({ action: 'insert', type: 'chart', anchor: 'p_001', position: 'after', chartType: 'doughnut', holeSize: bad }).success).toBe(false);
    }
    expect(chartHoleSizeSchemaRemedy('operations.0.holeSize: Number must be <= 90')).toContain('1 to 90');
    expect(chartHoleSizeSchemaRemedy('operations.0.chartType: Invalid enum value')).toBeUndefined();
  });

  it('refuses holeSize on non-doughnut types and preserves an authored hole on a structural update', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    expect(() => executeChart({ action: 'chart', chartType: 'pie', holeSize: 75 }, ast, parser)).toThrow(/holeSize is only supported for doughnut/);

    const foreign = await insertChart('doughnut', 'foreign_struct');
    foreign.parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      chartXml(foreign.parser).replace('<c:holeSize val="50"/>', '<c:holeSize val="75"/>')
    );
    executeChart({
      action: 'chart', operation: 'update', target: 'foreign_struct',
      series: [
        { name: 'ARR', values: [99, 20, 30] },
        { name: 'Costs', values: [5, 7, 9] },
      ],
    }, foreign.ast, foreign.parser);
    const after = chartXml(foreign.parser);
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="75"/>');
    expect(after).not.toContain('<c:bar3DChart>');
  });

  it('refuses multi-plot parts containing a doughnut before any mutation', async () => {
    for (const seriesCount of [2, 1]) {
      const { parser, ast, id } = await insertChart('doughnut', `combo_${seriesCount}`);
      const pieXml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
        .generate2DChartXml({ chartType: 'pie', title: 'Other', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Other', values: [7, 8, 9] }] });
      const piePlot = /<c:pieChart>[\s\S]*?<\/c:pieChart>/.exec(pieXml)![0];
      const combo = chartXml(parser).replace('</c:plotArea>', `${piePlot}</c:plotArea>`);
      parser.getZipAdapter().writeText('word/charts/chart1.xml', combo);
      const version = chartNode(ast, id).version;
      const series = seriesCount === 2
        ? [{ name: 'ARR', values: [99, 20, 30] }, { name: 'Other', values: [2, 8, 9] }]
        : [{ name: 'ARR', values: [99, 20, 30] }];
      let message = '';
      try {
        executeChart({ action: 'chart', operation: 'update', target: id, series }, ast, parser);
      } catch (err) {
        message = (err as Error).message;
      }
      expect(message).toContain('multi-plot chart part');
      expect(message).toContain('pieChart');
      expect(message).toContain('doughnutChart');
      expect(message).toContain('word/charts/chart1.xml');
      expect(chartXml(parser)).toBe(combo);
      expect(chartNode(ast, id).version).toBe(version);
    }
  });

  it('refuses same-tag duplicate plots (two doughnuts, bar plus bar)', async () => {
    for (const spec of [
      { chartType: 'doughnut' as const, tag: 'doughnutChart' },
      { chartType: 'column' as const, tag: 'barChart' },
    ]) {
      const { parser, ast, id } = await insertChart(spec.chartType, `dup_${spec.tag}`);
      const xml = chartXml(parser);
      const block = new RegExp(`<c:${spec.tag}>[\\s\\S]*?</c:${spec.tag}>`).exec(xml)![0];
      const duplicate = xml.replace('</c:plotArea>', `${block}</c:plotArea>`);
      parser.getZipAdapter().writeText('word/charts/chart1.xml', duplicate);
      const version = chartNode(ast, id).version;
      let message = '';
      try {
        executeChart({ action: 'chart', operation: 'update', target: id, series: [{ name: 'ARR', values: [99, 20, 30] }] }, ast, parser);
      } catch (err) {
        message = (err as Error).message;
      }
      expect(message).toContain('multi-plot chart part');
      expect(message).toContain(`${spec.tag}, ${spec.tag}`);
      expect(message).toContain('word/charts/chart1.xml');
      expect(chartXml(parser)).toBe(duplicate);
      expect(chartNode(ast, id).version).toBe(version);
    }
  });

  it('matches plot elements by local name across namespace prefixes', async () => {
    const NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

    // The doughnut plot itself may carry any prefix; it must stay doughnut.
    const single = await insertChart('doughnut', 'prefix_single');
    const singleXml = chartXml(single.parser);
    const prefixedDoughnut = singleXml
      .replace('<c:doughnutChart>', `<x:doughnutChart xmlns:x="${NS}">`)
      .replace('</c:doughnutChart>', '</x:doughnutChart>');
    single.parser.getZipAdapter().writeText('word/charts/chart1.xml', prefixedDoughnut);
    expect(detectChartKindFromXml(prefixedDoughnut)).toEqual({ is3D: false, chartType: 'doughnut' });
    expect(summarize2DChartXml(prefixedDoughnut).chartType).toBe('doughnut');

    const reopened = new DocxParser(single.parser.rezip());
    const ast2 = await reopened.getAST();
    const node = ast2.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('doughnut');
    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, ast2, reopened);
    expect(chartXml(reopened)).toMatch(/<x:doughnutChart[\s>]/);
    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }, { name: 'Costs', values: [5, 7, 9] }],
    }, ast2, reopened);
    const structural = chartXml(reopened);
    expect(structural).toContain('<c:doughnutChart>');
    expect(structural).not.toContain('<c:barChart>');
    expect(structural).toContain('<c:holeSize val="50"/>');

    // A prefixed sibling plot is still a second plot and must be refused.
    const comboBase = await insertChart('doughnut', 'prefix_combo');
    const barXml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'bar', title: 'Other', categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'Other', values: [7, 8, 9] }] });
    const barPlot = /<c:barChart>[\s\S]*?<\/c:barChart>/.exec(barXml)![0]
      .replace('<c:barChart>', `<x:barChart xmlns:x="${NS}">`)
      .replace('</c:barChart>', '</x:barChart>');
    const combo = chartXml(comboBase.parser).replace('</c:plotArea>', `${barPlot}</c:plotArea>`);
    comboBase.parser.getZipAdapter().writeText('word/charts/chart1.xml', combo);
    expect((await validateDocxBuffer(comboBase.parser.rezip())).ok).toBe(true);
    for (const seriesCount of [2, 1]) {
      const version = chartNode(comboBase.ast, comboBase.id).version;
      const series = seriesCount === 2
        ? [{ name: 'ARR', values: [99, 20, 30] }, { name: 'Other', values: [2, 8, 9] }]
        : [{ name: 'ARR', values: [99, 20, 30] }];
      let message = '';
      try {
        executeChart({ action: 'chart', operation: 'update', target: comboBase.id, series }, comboBase.ast, comboBase.parser);
      } catch (err) {
        message = (err as Error).message;
      }
      expect(message).toContain('multi-plot chart part');
      expect(message).toContain('barChart');
      expect(message).toContain('doughnutChart');
      expect(message).toContain('word/charts/chart1.xml');
      expect(chartXml(comboBase.parser)).toBe(combo);
      expect(chartNode(comboBase.ast, comboBase.id).version).toBe(version);
    }

    // The same two-plot shape with the doughnut (not the bar) prefixed.
    const swappedBase = await insertChart('doughnut', 'prefix_combo_swap');
    const swappedXml = chartXml(swappedBase.parser);
    const swappedPlot = /<c:doughnutChart>[\s\S]*?<\/c:doughnutChart>/.exec(swappedXml)![0];
    const plainBarPlot = /<c:barChart>[\s\S]*?<\/c:barChart>/.exec(barXml)![0];
    const swapped = swappedXml
      .replace(swappedPlot, swappedPlot
        .replace('<c:doughnutChart>', `<x:doughnutChart xmlns:x="${NS}">`)
        .replace('</c:doughnutChart>', '</x:doughnutChart>'))
      .replace('</c:plotArea>', `${plainBarPlot}</c:plotArea>`);
    swappedBase.parser.getZipAdapter().writeText('word/charts/chart1.xml', swapped);
    expect((await validateDocxBuffer(swappedBase.parser.rezip())).ok).toBe(true);
    let swappedMessage = '';
    try {
      executeChart({
        action: 'chart', operation: 'update', target: swappedBase.id,
        series: [{ name: 'ARR', values: [99, 20, 30] }],
      }, swappedBase.ast, swappedBase.parser);
    } catch (err) {
      swappedMessage = (err as Error).message;
    }
    expect(swappedMessage).toContain('multi-plot chart part');
    expect(swappedMessage).toContain('barChart, doughnutChart');
    expect(swappedMessage).toContain('word/charts/chart1.xml');
    expect(chartXml(swappedBase.parser)).toBe(swapped);

    // Three plots with three different prefixes count as three occurrences.
    const threeBase = await insertChart('doughnut', 'prefix_three');
    const threeXml = chartXml(threeBase.parser);
    const doughnutPlot = /<c:doughnutChart>[\s\S]*?<\/c:doughnutChart>/.exec(threeXml)![0];
    const withPrefix = (prefix: string) => doughnutPlot
      .replace('<c:doughnutChart>', `<${prefix}:doughnutChart xmlns:${prefix}="${NS}">`)
      .replace('</c:doughnutChart>', `</${prefix}:doughnutChart>`);
    const three = threeXml.replace(doughnutPlot, withPrefix('x') + withPrefix('y') + withPrefix('z'));
    threeBase.parser.getZipAdapter().writeText('word/charts/chart1.xml', three);
    expect((await validateDocxBuffer(threeBase.parser.rezip())).ok).toBe(true);
    let threeMessage = '';
    try {
      executeChart({
        action: 'chart', operation: 'update', target: threeBase.id,
        series: [{ name: 'ARR', values: [99, 20, 30] }],
      }, threeBase.ast, threeBase.parser);
    } catch (err) {
      threeMessage = (err as Error).message;
    }
    expect(threeMessage).toContain('multi-plot chart part');
    expect(threeMessage).toContain('doughnutChart, doughnutChart, doughnutChart');
    expect(threeMessage).toContain('word/charts/chart1.xml');
    expect(chartXml(threeBase.parser)).toBe(three);
  });

  it('ignores foreign-namespace plot elements nested in c:extLst', async () => {
    const pie = await insertChart('pie', 'vendor_ext_pie');
    const vendor = '  <c:extLst><c:ext uri="{BBE1A952-AA13-448E-AADC-164F8A28A991}">' +
      '<v:doughnutChart xmlns:v="urn:example:vendor"><v:ser><v:idx val="0"/></v:ser></v:doughnutChart>' +
      '</c:ext></c:extLst>';
    const xml = chartXml(pie.parser).replace('</c:chartSpace>', `${vendor}\n</c:chartSpace>`);
    pie.parser.getZipAdapter().writeText('word/charts/chart1.xml', xml);
    expect((await validateDocxBuffer(pie.parser.rezip())).ok).toBe(true);
    expect(detectChartKindFromXml(xml)).toEqual({ is3D: false, chartType: 'pie' });
    expect(summarize2DChartXml(xml).chartType).toBe('pie');

    const reopened = new DocxParser(pie.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('pie');
    expect(renderChartText(node)).toContain('[Chart pie]');

    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, ast, reopened);
    const after = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:pieChart>');
    expect(after).not.toMatch(/<c:doughnutChart[\s>]/);
    expect(after).toContain('<c:v>99</c:v>');
    expect(after).toContain('<v:doughnutChart');
    expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
  });

  it('does not treat a foreign-namespace bar plot in c:extLst as a second plot', async () => {
    const base = await insertChart('doughnut', 'vendor_ext_doughnut', { holeSize: 75 });
    const vendor = '  <c:extLst><c:ext uri="{BBE1A952-AA13-448E-AADC-164F8A28A991}">' +
      '<v:barChart xmlns:v="urn:example:vendor"><v:ser><v:idx val="0"/></v:ser></v:barChart>' +
      '</c:ext></c:extLst>';
    const xml = chartXml(base.parser).replace('</c:chartSpace>', `${vendor}\n</c:chartSpace>`);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', xml);
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);

    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('doughnut');
    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      holeSize: 60,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, ast, reopened);
    const after = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:doughnutChart>');
    expect(after).toContain('<c:holeSize val="60"/>');
    expect(after).toContain('<c:v>99</c:v>');
    expect(after).toContain('<v:barChart');
    expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
  });

  it('merges holeSize into a prefixed doughnut plot', async () => {
    const NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    const base = await insertChart('doughnut', 'prefix_hole', { holeSize: 75 });
    const xml = chartXml(base.parser)
      .replace('<c:doughnutChart>', `<x:doughnutChart xmlns:x="${NS}">`)
      .replace('</c:doughnutChart>', '</x:doughnutChart>');
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', xml);
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);

    const reopened = new DocxParser(base.parser.rezip());
    const ast = await reopened.getAST();
    const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('doughnut');
    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      holeSize: 60,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, ast, reopened);
    const after = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toMatch(/<x:doughnutChart[\s>]/);
    expect(after).toContain('<c:holeSize val="60"/>');
    expect(after).toContain('<c:v>99</c:v>');
    expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
  });

  it('reads and merges default-namespace and fully aliased chart parts', async () => {
    for (const spec of [
      { name: 'default', tag: 'holeSize', transform: toDefaultNamespace },
      { name: 'xalias', tag: 'x:doughnutChart', transform: toXAlias },
    ]) {
      const base = await insertChart('doughnut', `ns_${spec.name}`, { holeSize: 75 });
      const xml = spec.transform(chartXml(base.parser));
      base.parser.getZipAdapter().writeText('word/charts/chart1.xml', xml);
      expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);
      expect(detectChartKindFromXml(xml)).toEqual({ is3D: false, chartType: 'doughnut' });
      const summary = summarize2DChartXml(xml);
      expect(summary.chartType).toBe('doughnut');
      expect(summary.title).toBe(DEMO.title);
      expect(summary.categories).toEqual(DEMO.categories);
      expect(summary.series[0]?.values).toEqual([10, 20, 30]);

      const reopened = new DocxParser(base.parser.rezip());
      const ast = await reopened.getAST();
      const node = ast.sections[0].nodes.find((n) => n.type === 'chart')!;
      const data = node.metadata.chartData as { title?: string; categories?: string[]; series?: Array<{ values?: number[] }>; holeSize?: number };
      expect(node.metadata.chartType).toBe('doughnut');
      expect(data.title).toBe(DEMO.title);
      expect(data.categories).toEqual(DEMO.categories);
      expect(data.series?.[0]?.values).toEqual([10, 20, 30]);
      expect(data.holeSize).toBe(75);
      expect(renderChartText(node)).toContain('[Chart doughnut]');
      const row = new ToonGenerator().toon(ast, { fidelity: 'standard' }).rows.find((r) => r.id === node.id)!;
      expect(row.text).toContain('[Chart doughnut]');
      expect(row.flags).toContain('[CH:doughnut]');
      expect(row.flags).toMatch(/\[RID:rId\d+\]/);
      const html = convertAstToHtml(ast, `ns_${spec.name}`);
      expect(html).toContain('A 60 60 0');
      expect(html).not.toContain('A 40 40 0');

      const before = reopened.getZipAdapter().readText('word/charts/chart1.xml');
      executeChart({
        action: 'chart', operation: 'update', target: node.id,
        series: [{ name: 'ARR', values: [99, 20, 30] }],
      }, ast, reopened);
      const after = reopened.getZipAdapter().readText('word/charts/chart1.xml');
      expect(after.replace('>99<', '>10<')).toBe(before);
      expect(after).toMatch(/<[^>]*holeSize val="75"/);
      expect(after).toContain(spec.tag);
      expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
    }
  });

  it('resolves numeric character references in chart namespace declarations', async () => {
    const HEX_REF_NS = 'http://schemas.openxmlformats.org/drawingml/2006/char&#x74;';
    const DEC_REF_NS = 'http://schemas.openxmlformats.org/drawingml/2006/char&#116;';

    // A doughnut plot whose alias is declared with a hexadecimal reference.
    const single = await insertChart('doughnut', 'numeric_single', { holeSize: 75 });
    const aliased = chartXml(single.parser)
      .replace('<c:doughnutChart>', `<x:doughnutChart xmlns:x="${HEX_REF_NS}">`)
      .replace('</c:doughnutChart>', '</x:doughnutChart>');
    single.parser.getZipAdapter().writeText('word/charts/chart1.xml', aliased);
    expect((await validateDocxBuffer(single.parser.rezip())).ok).toBe(true);
    expect(detectChartKindFromXml(aliased)).toEqual({ is3D: false, chartType: 'doughnut' });
    expect(summarize2DChartXml(aliased).chartType).toBe('doughnut');

    const reopenedSingle = new DocxParser(single.parser.rezip());
    const astSingle = await reopenedSingle.getAST();
    const nodeSingle = astSingle.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(nodeSingle.metadata.chartType).toBe('doughnut');
    expect(renderChartText(nodeSingle)).toContain('[Chart doughnut]');
    expect(convertAstToHtml(astSingle, 'numeric_single')).toContain('A 60 60 0');
    executeChart({
      action: 'chart', operation: 'update', target: nodeSingle.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, astSingle, reopenedSingle);
    const afterSingle = reopenedSingle.getZipAdapter().readText('word/charts/chart1.xml');
    expectSingleValueSpan(aliased, afterSingle);
    expect(afterSingle).toMatch(/<x:doughnutChart[\s>]/);
    expect(afterSingle).toContain('<c:holeSize val="75"/>');

    // The default namespace declared with a decimal reference.
    const base = await insertChart('doughnut', 'numeric_default', { holeSize: 75 });
    const defaultXml = toDefaultNamespace(chartXml(base.parser))
      .replace(`xmlns="${CHART_NS_URI}"`, `xmlns="${DEC_REF_NS}"`);
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', defaultXml);
    expect((await validateDocxBuffer(base.parser.rezip())).ok).toBe(true);
    expect(detectChartKindFromXml(defaultXml)).toEqual({ is3D: false, chartType: 'doughnut' });
    const reopenedDefault = new DocxParser(base.parser.rezip());
    const astDefault = await reopenedDefault.getAST();
    const nodeDefault = astDefault.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(nodeDefault.metadata.chartType).toBe('doughnut');
    executeChart({
      action: 'chart', operation: 'update', target: nodeDefault.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, astDefault, reopenedDefault);
    const afterDefault = reopenedDefault.getZipAdapter().readText('word/charts/chart1.xml');
    expect(afterDefault.replace('>99<', '>10<')).toBe(defaultXml);
    expect(afterDefault).toContain('<holeSize val="75"/>');
    expect(afterDefault).not.toContain('<c:');

    // A sibling plot whose alias uses a reference is still a second plot.
    const comboBase = await insertChart('doughnut', 'numeric_combo');
    const barXml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
      .generate2DChartXml({ chartType: 'bar', title: 'Other', categories: ['Q1', 'Q2'], series: [{ name: 'Other', values: [3, 4] }] });
    const barPlot = /<c:barChart>[\s\S]*?<\/c:barChart>/.exec(barXml)![0]
      .replace('<c:barChart>', `<x:barChart xmlns:x="${HEX_REF_NS}">`)
      .replace('</c:barChart>', '</x:barChart>');
    const combo = chartXml(comboBase.parser).replace('</c:plotArea>', `${barPlot}</c:plotArea>`);
    comboBase.parser.getZipAdapter().writeText('word/charts/chart1.xml', combo);
    expect((await validateDocxBuffer(comboBase.parser.rezip())).ok).toBe(true);
    for (const series of [
      [{ name: 'ARR', values: [99, 20, 30] }, { name: 'Other', values: [77, 88] }],
      [{ name: 'ARR', values: [99, 20, 30] }],
    ]) {
      const version = chartNode(comboBase.ast, comboBase.id).version;
      let message = '';
      try {
        executeChart({ action: 'chart', operation: 'update', target: comboBase.id, series }, comboBase.ast, comboBase.parser);
      } catch (err) {
        message = (err as Error).message;
      }
      expect(message).toContain('multi-plot chart part');
      expect(message).toContain('barChart, doughnutChart');
      expect(message).toContain('word/charts/chart1.xml');
      expect(chartXml(comboBase.parser)).toBe(combo);
      expect(chartNode(comboBase.ast, comboBase.id).version).toBe(version);
    }
  });

  it('keeps a doughnut plus view3D as doughnut through read, HTML and updates', async () => {
    const { parser } = await insertChart('doughnut', 'v3d');
    const withView = chartXml(parser)
      .replace('<c:plotArea>', '<c:view3D><c:rotX val="30"/></c:view3D>\n    <c:plotArea>')
      .replace('<c:holeSize val="50"/>', '<c:holeSize val="75"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', withView);
    expect(detectChartKindFromXml(withView)).toEqual({ is3D: false, chartType: 'doughnut' });

    const reopened = new DocxParser(parser.rezip());
    const ast2 = await reopened.getAST();
    const node = ast2.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('doughnut');
    expect(renderChartText(node)).toContain('[Chart doughnut]');
    expect(convertAstToHtml(ast2, 'v3d')).toContain('A 60 60 0');

    const before = withView;
    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }],
    }, ast2, reopened);
    const afterValues = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expectSingleValueSpan(before, afterValues);
    expect(afterValues).toContain('<c:doughnutChart>');
    expect(afterValues).toContain('<c:holeSize val="75"/>');

    executeChart({
      action: 'chart', operation: 'update', target: node.id,
      series: [{ name: 'ARR', values: [99, 20, 30] }, { name: 'Costs', values: [5, 7, 9] }],
    }, ast2, reopened);
    const afterStruct = reopened.getZipAdapter().readText('word/charts/chart1.xml');
    expect(afterStruct).toContain('<c:doughnutChart>');
    expect(afterStruct).toContain('<c:holeSize val="75"/>');
    expect(afterStruct).not.toContain('<c:bar3DChart>');
    expect((await validateDocxBuffer(reopened.rezip())).ok).toBe(true);
  });

  it('HTML uses the authored holeSize and pins the single-slice ring', async () => {
    const authored = await insertChart('doughnut', 'hole_html');
    const authoredXml = chartXml(authored.parser).replace('<c:holeSize val="50"/>', '<c:holeSize val="75"/>');
    authored.parser.getZipAdapter().writeText('word/charts/chart1.xml', authoredXml);
    const reopened = new DocxParser(authored.parser.rezip());
    const html = convertAstToHtml(await reopened.getAST(), 'hole_html');
    expect(html).toContain('A 60 60 0');
    expect(html).not.toContain('A 40 40 0');

    const single = await insertChart('doughnut', 'single_ring', { categories: ['Only'], series: [{ name: 'Share', values: [42] }] });
    const singleHtml = convertAstToHtml(single.ast, 'single');
    expect(singleHtml).toContain('fill-rule="evenodd"');
    expect(singleHtml).toContain('a 40 40');
    expect(singleHtml).not.toContain('M 170 120 L');
  });

  it('names the offending tag and path for a self-closed doughnut plot', async () => {
    const { parser, ast, id } = await insertChart('doughnut', 'empty_plot');
    const before = chartXml(parser).replace(/<c:doughnutChart>[\s\S]*?<\/c:doughnutChart>/, '<c:doughnutChart/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', before);
    let message = '';
    try {
      executeChart({ action: 'chart', operation: 'update', target: id, series: [{ name: 'ARR', values: [99, 20, 30] }] }, ast, parser);
    } catch (err) {
      message = (err as Error).message;
    }
    expect(message).toContain('<c:doughnutChart/>');
    expect(message).toContain('word/charts/chart1.xml');
    expect(chartXml(parser)).toBe(before);
  });

});

function nodeChartData(ast: DocumentAST, id: string): { chartType?: string } {
  return chartNode(ast, id).metadata.chartData as { chartType?: string };
}
