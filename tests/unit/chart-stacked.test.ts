import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  Chart2DManipulator,
  Chart2DType,
  Chart3DManipulator,
  DocxParser,
  DocumentAST,
  BAR_2D_VARIANTS,
  OperationSchema,
  chartTypeSchemaRemedy,
  SUPPORTED_CHART_2D_AND_3D_TYPES,
  ToonGenerator,
  convertAstToHtml,
  detectBar2DChartTypeFromXml,
  detectChartKindFromXml,
  executeChart,
  renderChartText,
  summarize2DChartXml,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

/**
 * Stacked and percent-stacked column/bar pins: writer tags, reader detection,
 * insert path, in-place update routing (chartType equality keeps the merge
 * path), and OOXML validity. Doughnut/scatter stay out of this family.
 */

const DEMO = {
  title: 'Revenue by quarter',
  categories: ['Q1', 'Q2', 'Q3'],
  series: [
    { name: 'ARR', values: [10, 20, 30], color: '0066CC' },
    { name: 'Costs', values: [5, 7, 9], color: '10B981' },
  ],
};

const VARIANT_TYPES = [
  'columnStacked',
  'barStacked',
  'columnPercentStacked',
  'barPercentStacked',
] as const;

type VariantType = (typeof VARIANT_TYPES)[number];

const EXPECTED: Record<VariantType, { dir: 'col' | 'bar'; grouping: 'stacked' | 'percentStacked'; percent: boolean }> = {
  columnStacked: { dir: 'col', grouping: 'stacked', percent: false },
  barStacked: { dir: 'bar', grouping: 'stacked', percent: false },
  columnPercentStacked: { dir: 'col', grouping: 'percentStacked', percent: true },
  barPercentStacked: { dir: 'bar', grouping: 'percentStacked', percent: true },
};

function chartXmlOf(buf: Buffer, name = 'chart1.xml'): string {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  if (!entry) throw new Error(`${name} missing from rezipped docx`);
  return entry.getData().toString('utf8');
}

function chartNode(ast: DocumentAST, id: string) {
  return ast.sections[0].nodes.find((n) => n.id === id)!;
}

async function insertVariant(chartType: Chart2DType, id = `chart_${chartType}`) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  executeChart({ action: 'chart', chartType, id, ...DEMO }, ast, parser);
  const buf = parser.rezip();
  return { parser, ast, id, buf, xml: chartXmlOf(buf) };
}

describe('stacked and percent-stacked column/bar charts', () => {
  it.each(VARIANT_TYPES)('%s writes barDir, grouping and full overlap', (chartType) => {
    const expected = EXPECTED[chartType];
    const xml = new Chart2DManipulator({ sections: [], relationships: {} } as never).generate2DChartXml({
      chartType,
      ...DEMO,
    });
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain(`<c:barDir val="${expected.dir}"/>`);
    expect(xml).toContain(`<c:grouping val="${expected.grouping}"/>`);
    expect(xml).toContain(`<c:gapWidth val="${expected.dir === 'bar' ? 120 : 150}"/>`);
    expect(xml).toContain('<c:overlap val="100"/>');
    expect(xml).not.toContain('<c:overlap val="-30"/>');
    expect(xml).not.toContain('3DChart');
    expect(xml).not.toContain('<c:view3D>');
    if (expected.percent) {
      expect(xml).toContain('<c:numFmt formatCode="0%" sourceLinked="0"/>');
      expect(xml).not.toContain('formatCode="General"');
    } else {
      expect(xml).toContain('<c:numFmt formatCode="General" sourceLinked="1"/>');
      expect(xml).not.toContain('formatCode="0%"');
    }
  });

  it.each(VARIANT_TYPES)('%s inserts a valid chart part with matching metadata', async (chartType) => {
    const { parser, ast, id, buf, xml } = await insertVariant(chartType);
    const expected = EXPECTED[chartType];
    expect(xml).toContain(`<c:grouping val="${expected.grouping}"/>`);
    expect(xml).toContain('<c:overlap val="100"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe(chartType);
    expect(chartNode(ast, id).type).toBe('chart');
    const docx = await validateDocxBuffer(buf);
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
    const ooxml = await parser.validateOOXML();
    expect(ooxml.ok).toBe(true);
    expect(ooxml.errors).toEqual([]);
  });

  it.each(VARIANT_TYPES)('%s is detected from its part and surfaced by the reader', async (chartType) => {
    const { xml } = await insertVariant(chartType);
    expect(detectChartKindFromXml(xml)).toEqual({ is3D: false, chartType });
    expect(summarize2DChartXml(xml).chartType).toBe(chartType);
  });

  it('operation schema enumerates every supported chart type and no later-batch family', () => {
    for (const chartType of SUPPORTED_CHART_2D_AND_3D_TYPES) {
      const parsed = OperationSchema.safeParse({
        action: 'chart',
        operation: 'insert',
        anchor: 'p_001',
        chartType,
        categories: ['Q1'],
        series: [{ name: 'S', values: [1] }],
      });
      expect(parsed.success).toBe(true);
    }
    for (const chartType of ['stackedColumn', 'columnStacked100', 'bubble']) {
      expect(
        OperationSchema.safeParse({
          action: 'chart',
          operation: 'insert',
          anchor: 'p_001',
          chartType,
          categories: ['Q1'],
          series: [{ name: 'S', values: [1] }],
        }).success
      ).toBe(false);
    }
    const legacy = (chartType: string) => OperationSchema.safeParse({
      action: 'insert',
      type: 'chart',
      anchor: 'p_001',
      position: 'after',
      chartType,
      categories: ['Q1'],
      series: [{ name: 'S', values: [1] }],
    }).success;
    expect(legacy('columnStacked')).toBe(true);
    expect(legacy('barPercentStacked')).toBe(true);
    expect(legacy('doughnut')).toBe(true);
    expect(legacy('scatter')).toBe(true);
    expect(chartTypeSchemaRemedy('operations.0.chartType: Invalid enum value')).toContain('omit chartType');
    expect(chartTypeSchemaRemedy('operations.0.action: Invalid input')).toBeUndefined();
  });

  it('writer and reader consume one variant matrix', () => {
    for (const [chartType, variant] of Object.entries(BAR_2D_VARIANTS)) {
      const xml = new Chart2DManipulator({ sections: [], relationships: {} } as never)
        .generate2DChartXml({ chartType: chartType as Chart2DType, ...DEMO });
      expect(xml).toContain(`<c:barDir val="${variant.barDir}"/>`);
      expect(xml).toContain(`<c:grouping val="${variant.grouping}"/>`);
      expect(detectBar2DChartTypeFromXml(xml)).toBe(chartType);
      expect(detectChartKindFromXml(xml).chartType).toBe(chartType);
    }
  });

  it('TOON surfaces the full chart type so inspect output can round-trip', async () => {
    for (const [chartType] of Object.entries(BAR_2D_VARIANTS)) {
      const { ast, id } = await insertVariant(chartType as Chart2DType);
      const node = chartNode(ast, id);
      expect(renderChartText(node)).toContain(`[Chart ${chartType}]`);
      const row = new ToonGenerator().toon(ast, { fidelity: 'standard' }).rows.find((r) => r.id === id)!;
      expect(row.flags).toContain(`[CH:${chartType}]`);
    }
  });

  it('HTML export stacks the variants instead of grouping absolute bars', async () => {
    const clustered = await insertVariant('column');
    const clusteredHtml = convertAstToHtml(clustered.ast, 'clustered');
    expect(clusteredHtml).toContain('>30<');
    expect(clusteredHtml).not.toContain('>39<');

    const stacked = await insertVariant('columnStacked');
    const stackedHtml = convertAstToHtml(stacked.ast, 'stacked');
    expect(stackedHtml).toContain('>39<');

    const percent = await insertVariant('columnPercentStacked');
    const percentHtml = convertAstToHtml(percent.ast, 'percent');
    expect(percentHtml).toContain('>100%<');
    expect(percentHtml).toContain('>67%<');
    expect(percentHtml).toContain('>33%<');
  });

  it('reads a clustered bar part as the legacy names (no grouping regression)', () => {
    const m = new Chart2DManipulator({ sections: [], relationships: {} } as never);
    expect(detectChartKindFromXml(m.generate2DChartXml({ chartType: 'column', ...DEMO }))).toEqual({
      is3D: false,
      chartType: 'column',
    });
    expect(detectChartKindFromXml(m.generate2DChartXml({ chartType: 'bar', ...DEMO }))).toEqual({
      is3D: false,
      chartType: 'bar',
    });
  });

  it('scopes grouping detection to c:barChart (3D bar parts stay 3D)', () => {
    const xml3d = new Chart3DManipulator({ sections: [], relationships: {} } as never)
      .generate3DChartXml({ chartType: 'bar3D', ...DEMO })
      .replace('<c:grouping val="clustered"/>', '<c:grouping val="stacked"/>');
    expect(detectChartKindFromXml(xml3d)).toEqual({ is3D: true, chartType: 'bar3D' });
  });

  it.each(VARIANT_TYPES)('%s keeps grouping bytes on an explicit same-type update', async (chartType) => {
    const { parser, ast, id } = await insertVariant(chartType);
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        chartType,
        series: [
          { name: 'ARR', values: [99, 20, 30] },
          { name: 'Costs', values: [5, 7, 9] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const oldSpan = '<c:v>10</c:v>';
    const at = before.indexOf(oldSpan);
    expect(at).toBeGreaterThan(0);
    expect(after).toBe(before.slice(0, at) + '<c:v>99</c:v>' + before.slice(at + oldSpan.length));
    expect(after).toContain(`<c:grouping val="${EXPECTED[chartType].grouping}"/>`);
    expect(after).toContain('<c:overlap val="100"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe(chartType);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('updates a foreign stacked part through chartType equality without touching grouping bytes', async () => {
    const { parser, ast, id } = await insertVariant('column');
    const foreign = parser
      .getZipAdapter()
      .readText('word/charts/chart1.xml')
      .replace('<c:grouping val="clustered"/>', '<c:grouping val="percentStacked"/>')
      .replace('<c:overlap val="-30"/>', '<c:overlap val="100"/>')
      .replace('<c:numFmt formatCode="General" sourceLinked="1"/>', '<c:numFmt formatCode="0%" sourceLinked="0"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', foreign);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        chartType: 'columnPercentStacked',
        series: [
          { name: 'ARR', values: [99, 20, 30] },
          { name: 'Costs', values: [5, 7, 9] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const oldSpan = '<c:v>10</c:v>';
    const at = foreign.indexOf(oldSpan);
    expect(after).toBe(foreign.slice(0, at) + '<c:v>99</c:v>' + foreign.slice(at + oldSpan.length));
    expect(after).toContain('<c:grouping val="percentStacked"/>');
    expect(after).toContain('<c:overlap val="100"/>');
    expect(after).toContain('<c:numFmt formatCode="0%" sourceLinked="0"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe('columnPercentStacked');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('detects the existing part on load so update routing sees the stacked family', async () => {
    const { buf } = await insertVariant('columnPercentStacked', 'loaded_chart');
    const reopened = new DocxParser(buf);
    const ast = await reopened.getAST();
    const node = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'chart')!;
    expect(node.metadata.chartType).toBe('columnPercentStacked');
    expect((node.metadata.chartData as { chartType?: string }).chartType).toBe('columnPercentStacked');
    expect(node.type).toBe('chart');
  });

  it.each([
    ['columnStacked', 'column'],
    ['barStacked', 'bar'],
    ['columnPercentStacked', 'column'],
    ['barPercentStacked', 'bar'],
  ] as const)('base family on %s keeps grouping (%s request)', async (from, baseFamily) => {
    const { parser, ast, id } = await insertVariant(from);
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    executeChart({
      action: 'chart',
      operation: 'update',
      target: id,
      chartType: baseFamily,
      series: [
        { name: 'ARR', values: [99, 20, 30] },
        { name: 'Costs', values: [5, 7, 9] },
      ],
    }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const oldSpan = '<c:v>10</c:v>';
    const at = before.indexOf(oldSpan);
    expect(after).toBe(before.slice(0, at) + '<c:v>99</c:v>' + before.slice(at + oldSpan.length));
    expect(after).toContain(`<c:grouping val="${EXPECTED[from].grouping}"/>`);
    expect(after).toContain('<c:overlap val="100"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe(from);
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it.each(VARIANT_TYPES)('%s values-only update with no chartType merges in place', async (chartType) => {
    const { parser, ast, id } = await insertVariant(chartType);
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    executeChart({
      action: 'chart',
      operation: 'update',
      target: id,
      series: [
        { name: 'ARR', values: [99, 20, 30] },
        { name: 'Costs', values: [5, 7, 9] },
      ],
    }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const oldSpan = '<c:v>10</c:v>';
    const at = before.indexOf(oldSpan);
    expect(after).toBe(before.slice(0, at) + '<c:v>99</c:v>' + before.slice(at + oldSpan.length));
    expect(after).toContain(`<c:grouping val="${EXPECTED[chartType].grouping}"/>`);
    expect(after).toContain('<c:overlap val="100"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe(chartType);
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('a series-count change regenerates in the detected stacked family, not clustered', async () => {
    const { parser, ast, id } = await insertVariant('columnStacked');
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [10, 20, 30] },
          { name: 'Costs', values: [5, 7, 9] },
          { name: 'Margin', values: [1, 2, 3] },
        ],
      },
      ast,
      parser
    );
    const xml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(xml).toContain('<c:grouping val="stacked"/>');
    expect(xml).toContain('<c:overlap val="100"/>');
    expect(xml).toContain('<c:ser>');
    expect(chartNode(ast, id).metadata.chartType).toBe('columnStacked');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('changing stacked to percent-stacked regenerates with the percent axis', async () => {
    const { parser, ast, id } = await insertVariant('columnStacked');
    executeChart({ action: 'chart', operation: 'update', target: id, chartType: 'columnPercentStacked' }, ast, parser);
    const xml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(xml).toContain('<c:grouping val="percentStacked"/>');
    expect(xml).toContain('<c:numFmt formatCode="0%" sourceLinked="0"/>');
    expect(chartNode(ast, id).metadata.chartType).toBe('columnPercentStacked');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });
});
