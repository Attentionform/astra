import { describe, it, expect } from 'vitest';
import crypto from 'node:crypto';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  DocumentAST,
  OperationSchema,
  convertAstToHtml,
  executeChart,
  summarize2DChartXml,
  validateDocxBuffer,
  validateOperationStrict,
  type Chart2DSeries,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const SERIES = [
  { name: 'Revenue', values: [10, 20, 30], chartType: 'column' as const, color: '0066CC' },
  { name: 'Margin', values: [0.2, 0.3, 0.4], chartType: 'line' as const, axis: 'secondary' as const, color: '10B981' },
];

function chartXmlOf(buf: Buffer, name = 'chart1.xml'): string {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  if (!entry) throw new Error(`${name} missing from docx`);
  return entry.getData().toString('utf8');
}

function hash(buf: Buffer): string {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

async function insertCombo(id = 'combo', series: Array<Record<string, unknown>> = SERIES, opts: Record<string, unknown> = {}) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  executeChart(
    { action: 'chart', chartType: 'combo', id, title: 'Revenue and Margin', categories: ['Q1', 'Q2', 'Q3'], series: series as unknown as Chart2DSeries[], ...opts },
    ast,
    parser
  );
  const buf = parser.rezip();
  return { parser, ast, id, buf, xml: chartXmlOf(buf) };
}

function nodeById(ast: DocumentAST, id: string) {
  return ast.sections[0].nodes.find((n) => n.id === id)!;
}

describe('combo charts and secondary axes', () => {
  it('writes one barChart and one lineChart sharing primary axes, with a secondary line plot', async () => {
    const { parser, ast, id, xml } = await insertCombo();
    expect(xml).toContain('<c:barChart>');
    expect(xml).toContain('<c:lineChart>');
    expect(xml.match(/<c:barChart>/g)).toHaveLength(1);
    expect(xml.match(/<c:lineChart>/g)).toHaveLength(1);
    expect(xml.match(/<c:catAx>/g)).toHaveLength(2);
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2);
    expect(xml).toContain('<c:axId val="2001"/>');
    expect(xml).toContain('<c:axId val="2002"/>');
    expect(xml).toContain('<c:delete val="1"/>');
    expect(xml.match(/<c:ser>/g)).toHaveLength(2);
    expect(xml.match(/<c:idx val="0"\/>/g)).toHaveLength(1);
    expect(xml.match(/<c:idx val="1"\/>/g)).toHaveLength(1);
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      chartType: 'combo',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'Revenue', values: [10, 20, 30], chartType: 'column', axis: 'primary' },
        { name: 'Margin', values: [0.2, 0.3, 0.4], chartType: 'line', axis: 'secondary' },
      ],
    });

    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      chartType: 'combo',
      series: [
        { name: 'Revenue', chartType: 'column', axis: 'primary' },
        { name: 'Margin', chartType: 'line', axis: 'secondary' },
      ],
    });
    const summary = summarize2DChartXml(xml);
    expect(summary.chartType).toBe('combo');
    expect(summary.series.map((s) => [s.chartType, s.axis])).toEqual([
      ['column', 'primary'], ['line', 'secondary'],
    ]);
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('merges values across both plots by series idx and preserves unparsed styling', async () => {
    const { parser, ast, id, xml } = await insertCombo('combo_merge');
    const styled = xml.replace('<c:legendPos val="r"/>', '<c:legendPos val="t"/>')
      .replace('<c:barChart>', '<c:barChart><c:spPr><a:noFill/></c:spPr>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', styled);
    executeChart(
      {
        action: 'chart', operation: 'update', target: id,
        categories: ['H1', 'H2', 'H3'],
        series: [
          { name: 'Revenue', values: [11, 21, 31], chartType: 'column' },
          { name: 'Margin', values: [0.25, 0.35, 0.45], chartType: 'line', axis: 'secondary' },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:v>11</c:v>');
    expect(after).toContain('<c:v>0.25</c:v>');
    expect(after).not.toContain('<c:v>10</c:v>');
    expect(after).toContain('<c:legendPos val="t"/>');
    expect(after).toContain('<c:spPr><a:noFill/></c:spPr>');
    expect(after).toContain('<c:pt idx="0"><c:v>H1</c:v></c:pt>');
    const blocks = after.match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(blocks[0]).toContain('Sheet1!$B$2:$B$4');
    expect(blocks[1]).toContain('Sheet1!$C$2:$C$4');
    expect(nodeById(ast, id).version).toBe(2);
  });

  it('regenerates when a series moves between plot type or axis', async () => {
    const { parser, ast, id } = await insertCombo('combo_structural');
    executeChart(
      {
        action: 'chart', operation: 'update', target: id,
        style: { legend: { position: 'b' } },
        series: [
          { name: 'Revenue', values: [12, 22, 32], chartType: 'column' },
          { name: 'Margin', values: [0.22, 0.32, 0.42], chartType: 'line', axis: 'primary' },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after.match(/<c:valAx>/g)).toHaveLength(1);
    expect(after).toContain('<c:legendPos val="b"/>');
    expect(after).toContain('<c:v>12</c:v>');
    expect(after).toContain('<c:v>0.22</c:v>');
  });

  it('keeps a legal primary-only combo with one catAx and one valAx', async () => {
    const { xml } = await insertCombo('combo_primary', [
      { name: 'Revenue', values: [10, 20, 30], chartType: 'column' },
      { name: 'Margin', values: [0.2, 0.3, 0.4], chartType: 'line', axis: 'primary' },
    ]);
    expect(xml.match(/<c:catAx>/g)).toHaveLength(1);
    expect(xml.match(/<c:valAx>/g)).toHaveLength(1);
    expect(xml).not.toContain('<c:axId val="2001"/>');
  });

  it('accepts combo schema and rejects invalid series field values', () => {
    const base = {
      action: 'chart', chartType: 'combo', categories: ['Q1'],
      series: [
        { name: 'A', values: [1], chartType: 'column' },
        { name: 'B', values: [2], chartType: 'line', axis: 'secondary' },
      ],
    };
    expect(OperationSchema.safeParse(base).success).toBe(true);
    expect(OperationSchema.safeParse({
      ...base,
      series: [{ name: 'A', values: [1], chartType: 'area' }, { name: 'B', values: [2], chartType: 'line' }],
    }).success).toBe(false);
    expect(OperationSchema.safeParse({
      ...base,
      series: [{ name: 'A', values: [1], chartType: 'column', axis: 'second' }],
    }).success).toBe(false);
  });

  it('refuses combo payloads without a column or line before any insert', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const before = hash(parser.rezip());
    expect(() => executeChart({
      action: 'chart', chartType: 'combo', id: 'bad_combo',
      categories: ['Q1'], series: [
        { name: 'A', values: [1], chartType: 'column' },
        { name: 'B', values: [2], chartType: 'column' },
      ],
    }, ast, parser)).toThrow(/at least one line series/);
    expect(hash(parser.rezip())).toBe(before);
    expect(ast.sections[0].nodes.some((n) => n.id === 'bad_combo')).toBe(false);
  });

  it('refuses a secondary column and keeps the package unchanged', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const before = hash(parser.rezip());
    expect(() => executeChart({
      action: 'chart', chartType: 'combo', id: 'bad_secondary',
      categories: ['Q1'], series: [
        { name: 'A', values: [1], chartType: 'column', axis: 'secondary' },
        { name: 'B', values: [2], chartType: 'line' },
      ],
    }, ast, parser)).toThrow(/secondary axis/);
    expect(hash(parser.rezip())).toBe(before);
  });

  it('still refuses a foreign bar+line axis signature and writes nothing', async () => {
    const base = await insertCombo('combo_foreign');
    const foreign = base.xml.replace(
      /<c:lineChart>[\s\S]*?<\/c:lineChart>/,
      '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:axId val="1001"/><c:axId val="9999"/></c:lineChart>'
    );
    base.parser.getZipAdapter().writeText('word/charts/chart1.xml', foreign);
    const before = hash(base.parser.rezip());
    expect(() => executeChart(
      { action: 'chart', operation: 'update', target: base.id, series: [
        { name: 'Revenue', values: [11, 21, 31], chartType: 'column' },
        { name: 'Margin', values: [0.25, 0.35, 0.45], chartType: 'line', axis: 'secondary' },
      ] },
      base.ast,
      base.parser
    )).toThrow(/unsupported combo axis signature/);
    expect(hash(base.parser.rezip())).toBe(before);
  });

  it('renders combo bars and lines in HTML export', async () => {
    const { ast, id } = await insertCombo('combo_html');
    const html = convertAstToHtml(ast, id);
    expect(html).toContain('<rect');
    expect(html).toContain('<polyline');
    expect(html).toContain('Revenue');
    expect(html).toContain('Margin');
  });
});

describe('combo style integration', () => {
  it('applies data labels to both combo plots', async () => {
    const { xml } = await insertCombo('combo_labels', SERIES, {
      style: { dataLabels: { show: true, position: 'outEnd' } },
    });
    expect(xml.match(/<c:dLbls>/g)).toHaveLength(2);
    expect(xml.match(/<c:dLblPos val="outEnd"\/>/g)).toHaveLength(2);
  });

  it('applies a trendline to the selected combo series by global idx', async () => {
    const { xml } = await insertCombo('combo_trend', SERIES, {
      style: { trendline: { type: 'linear', series: [2] } },
    });
    const blocks = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(blocks[0]).not.toContain('<c:trendline>');
    expect(blocks[1]).toContain('<c:trendline>');
  });
});

describe('combo with both primary and secondary line plots', () => {
  it('writes two lineChart plots and updates all three series by idx', async () => {
    const series = [
      { name: 'Revenue', values: [10, 20], chartType: 'column' as const },
      { name: 'Margin', values: [0.2, 0.3], chartType: 'line' as const, axis: 'primary' as const },
      { name: 'Rate', values: [1, 2], chartType: 'line' as const, axis: 'secondary' as const },
    ];
    const { parser, ast, id, xml } = await insertCombo('combo_three', series);
    expect(xml.match(/<c:lineChart>/g)).toHaveLength(2);
    expect(xml.match(/<c:ser>/g)).toHaveLength(3);
    executeChart(
      {
        action: 'chart', operation: 'update', target: id,
        series: [
          { name: 'Revenue', values: [11, 21], chartType: 'column' },
          { name: 'Margin', values: [0.21, 0.31], chartType: 'line', axis: 'primary' },
          { name: 'Rate', values: [1.1, 2.1], chartType: 'line', axis: 'secondary' },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:v>11</c:v>');
    expect(after).toContain('<c:v>0.21</c:v>');
    expect(after).toContain('<c:v>1.1</c:v>');
    expect(after.match(/<c:lineChart>/g)).toHaveLength(2);
  });
});

const THREE_SERIES = [
  { name: 'Revenue', values: [10, 20], chartType: 'column' as const },
  { name: 'Margin', values: [0.2, 0.3], chartType: 'line' as const, axis: 'primary' as const },
  { name: 'Rate', values: [1, 2], chartType: 'line' as const, axis: 'secondary' as const },
];

describe('combo data updates preserve stored plot and axis assignment', () => {
  it('keeps the secondary axis and metadata when every series omits chartType and axis', async () => {
    const { parser, ast, id } = await insertCombo('combo_values_only');
    executeChart(
      {
        action: 'chart', operation: 'update', target: id,
        series: [
          { name: 'Revenue', values: [11, 21, 31] },
          { name: 'Margin', values: [0.25, 0.35, 0.45] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after.match(/<c:catAx>/g)).toHaveLength(2);
    expect(after.match(/<c:valAx>/g)).toHaveLength(2);
    expect(after).toContain('<c:axId val="2002"/>');
    expect(after).toContain('<c:v>0.25</c:v>');
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      series: [
        { name: 'Revenue', chartType: 'column', axis: 'primary' },
        { name: 'Margin', chartType: 'line', axis: 'secondary' },
      ],
    });
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      series: [
        { name: 'Revenue', chartType: 'column', axis: 'primary' },
        { name: 'Margin', chartType: 'line', axis: 'secondary' },
      ],
    });
  });

  it('keeps a line series a line when only its values change', async () => {
    const { parser, ast, id } = await insertCombo('combo_line_values');
    executeChart(
      {
        action: 'chart', operation: 'update', target: id,
        series: [
          { name: 'Revenue', values: [10, 20, 30], chartType: 'column' },
          { name: 'Margin', values: [0.5, 0.6, 0.7] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after.match(/<c:lineChart>/g)).toHaveLength(1);
    expect(after.match(/<c:valAx>/g)).toHaveLength(2);
    expect(after).toContain('<c:v>0.5</c:v>');
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      series: [{ chartType: 'column' }, { chartType: 'line', axis: 'secondary' }],
    });
  });

  it('retains secondary line plot on a values-only update when series have non-zero or non-standard idx values', async () => {
    const { parser, ast, id } = await insertCombo('combo_nonzero_idx', [
      { name: 'Margin', values: [0.2, 0.3], chartType: 'line', axis: 'primary' },
      { name: 'Revenue', values: [10, 20], chartType: 'column' },
      { name: 'Rate', values: [1, 2], chartType: 'line', axis: 'secondary' },
    ]);

    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'Margin', values: [0.22, 0.33] },
          { name: 'Revenue', values: [12, 22] },
          { name: 'Rate', values: [1.2, 2.2] },
        ],
      },
      ast,
      parser
    );

    const chartNode = nodeById(ast, id);
    expect(chartNode.metadata.chartData).toMatchObject({
      chartType: 'combo',
      series: [
        { name: 'Margin', values: [0.22, 0.33], chartType: 'line', axis: 'primary' },
        { name: 'Revenue', values: [12, 22], chartType: 'column', axis: 'primary' },
        { name: 'Rate', values: [1.2, 2.2], chartType: 'line', axis: 'secondary' },
      ],
    });

    const afterXml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(afterXml).toContain('<c:axId val="2001"/>');
    expect(afterXml).toContain('<c:axId val="2002"/>');
    expect(afterXml).toContain('<c:v>1.2</c:v>');

    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const reopenedChart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(reopenedChart.metadata.chartData).toMatchObject({
      series: [
        { name: 'Margin', chartType: 'line', axis: 'primary' },
        { name: 'Revenue', chartType: 'column', axis: 'primary' },
        { name: 'Rate', chartType: 'line', axis: 'secondary' },
      ],
    });
  });
});

describe('combo data labels and trendlines across plots', () => {
  it('applies data labels to all three supported plots and reads them back', async () => {
    const { parser, xml } = await insertCombo('combo_three_labels', THREE_SERIES, {
      style: { dataLabels: { show: true, position: 'outEnd' } },
    });
    expect(xml.match(/<c:lineChart>/g)).toHaveLength(2);
    expect(xml.match(/<c:dLbls>/g)).toHaveLength(3);
    expect(xml.match(/<c:dLblPos val="outEnd"\/>/g)).toHaveLength(3);
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: { dataLabels: { show: true, position: 'outEnd' } },
    });
  });

  it('does not report labels as chart-wide when only one combo plot carries them', async () => {
    const { parser, xml } = await insertCombo('combo_three_labels_partial', THREE_SERIES, {
      style: { dataLabels: { show: true } },
    });
    const segments = xml.split('<c:lineChart>');
    const onlyBarLabels = segments
      .map((segment, index) => index === 0 ? segment : segment.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, ''))
      .join('<c:lineChart>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', onlyBarLabels);
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    const style = (chart.metadata.chartData as Record<string, unknown>).style as Record<string, unknown> | undefined;
    expect(style?.dataLabels).toBeUndefined();
  });

  it('reads a trendline from the secondary line by its c:idx', async () => {
    const { parser } = await insertCombo('combo_trend_secondary', SERIES, {
      style: { trendline: { type: 'linear', series: [2] } },
    });
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: { trendline: { type: 'linear', series: [2] } },
    });
  });

  it('reads a first-plot trendline by c:idx instead of plot position', async () => {
    const interleaved = [
      { name: 'Margin', values: [0.2, 0.3], chartType: 'line' as const, axis: 'primary' as const },
      { name: 'Revenue', values: [10, 20], chartType: 'column' as const },
      { name: 'Rate', values: [1, 2], chartType: 'line' as const, axis: 'secondary' as const },
    ];
    const { parser } = await insertCombo('combo_trend_first_plot', interleaved, {
      style: { trendline: { type: 'linear', series: [2] } },
    });
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: { trendline: { type: 'linear', series: [2] } },
    });
  });

  it('keeps the previous trendline metadata and XML when distinct types coexist', async () => {
    const { parser, ast, id, xml } = await insertCombo('combo_trend_mixed', THREE_SERIES, {
      style: { trendline: { type: 'linear', series: [2] } },
    });
    const segments = xml.split('</c:ser>');
    segments[2] = segments[2].replace(
      '<c:cat>',
      '<c:trendline><c:trendlineType val="exponential"/></c:trendline><c:cat>'
    );
    parser.getZipAdapter().writeText('word/charts/chart1.xml', segments.join('</c:ser>'));
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { dataLabels: { show: true } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:trendlineType val="linear"/>');
    expect(after).toContain('<c:trendlineType val="exponential"/>');
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      style: { trendline: { type: 'linear', series: [2] } },
    });
  });
});

describe('combo axis styling and HTML scaling', () => {
  it('copies the value number format to the secondary value axis only', async () => {
    const { xml } = await insertCombo('combo_secondary_format', THREE_SERIES, {
      style: { axes: { value: { numberFormat: '0.0%', min: 0, max: 1, title: 'Share' } } },
    });
    const axes = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    const secondary = axes.find((block) => block.includes('<c:axId val="2002"/>'))!;
    expect(secondary).toContain('<c:numFmt formatCode="0.0%" sourceLinked="0"/>');
    expect(secondary).not.toContain('<c:majorGridlines');
    expect(secondary).not.toContain('<c:title>');
    expect(secondary).not.toContain('<c:min ');
    expect(secondary).not.toContain('<c:max ');
    const primary = axes.find((block) => block.includes('<c:axId val="1002"/>'))!;
    expect(primary).toContain('<c:numFmt formatCode="0.0%" sourceLinked="0"/>');
  });

  it('scales combo bars and primary lines together in HTML export', async () => {
    const { ast, id } = await insertCombo('combo_html_scale', [
      { name: 'Bars', values: [1, 2], chartType: 'column' as const },
      { name: 'Line', values: [10, 20], chartType: 'line' as const, axis: 'primary' as const },
    ], { categories: ['Q1', 'Q2'] });
    const html = convertAstToHtml(ast, id);
    const polyline = /<polyline points="([^"]+)"/.exec(html)![1];
    const ys = polyline.split(' ').map((point) => Number(point.split(',')[1]));
    expect(ys.length).toBe(2);
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(40);
    expect(Math.max(...ys)).toBeLessThanOrEqual(180);
  });

  it('isolates primary axis scale from secondary line series in combo chart html export', async () => {
    const { ast, id } = await insertCombo('combo_html_secondary_scale', [
      { name: 'Revenue', values: [10, 20], chartType: 'column' as const, axis: 'primary' as const },
      { name: 'Margin', values: [1000, 2000], chartType: 'line' as const, axis: 'secondary' as const },
    ], { categories: ['Q1', 'Q2'] });
    const html = convertAstToHtml(ast, id);

    // Primary Y-axis tick labels should reflect primary scale [0, 20], not secondary scale [0, 2000]
    const tickMatches = [...html.matchAll(/<text x="42" [^>]*>([^<]+)<\/text>/g)].map((m) => m[1].trim());
    expect(tickMatches).toEqual(['0', '5', '10', '15', '20']);

    // Bar heights should correspond to primary range 20 (plotHeight = 140)
    // Value 10 -> height 70, Value 20 -> height 140
    const barMatches = [...html.matchAll(/<rect x="[^"]+" y="([^"]+)" width="([^"]+)" height="([^"]+)" [^>]*>\s*<title>([^:]+): ([^<]+)<\/title>/g)]
      .filter((m) => parseFloat(m[2]) !== 10 && parseFloat(m[3]) !== 10)
      .map((m) => ({
        y: parseFloat(m[1]),
        height: parseFloat(m[3]),
        series: m[4].trim(),
        value: parseFloat(m[5]),
      }));

    expect(barMatches).toHaveLength(2);
    expect(barMatches[0]).toMatchObject({ value: 10, height: 70, y: 110 });
    expect(barMatches[1]).toMatchObject({ value: 20, height: 140, y: 40 });

    // Secondary line should also still be rendered with its points
    expect(html).toContain('<polyline points="');
  });
});

describe('combo-only series selectors fail closed outside combo charts', () => {
  it('rejects chartType/axis on an explicit non-combo insert at the schema validator', () => {
    expect(() => validateOperationStrict({
      action: 'chart', operation: 'insert', chartType: 'column', id: 'bad_fields',
      categories: ['Q1'], series: [
        { name: 'A', values: [1], chartType: 'column' },
        { name: 'B', values: [2] },
      ],
    })).toThrow(/outside a combo chart/);
  });

  it('rejects the fields before insertion and leaves the package unchanged', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const before = hash(parser.rezip());
    expect(() => executeChart({
      action: 'chart', chartType: 'column', id: 'bad_fields_insert',
      categories: ['Q1'], series: [
        { name: 'A', values: [1], chartType: 'line' },
        { name: 'B', values: [2] },
      ],
    }, ast, parser)).toThrow(/not "combo"/);
    expect(hash(parser.rezip())).toBe(before);
    expect(ast.sections[0].nodes.some((n) => n.id === 'bad_fields_insert')).toBe(false);
  });

  it('rejects the fields on a non-combo update with zero mutation', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({
      action: 'chart', chartType: 'column', id: 'plain_update', categories: ['Q1'],
      series: [{ name: 'A', values: [1] }],
    }, ast, parser);
    const before = hash(parser.rezip());
    expect(() => executeChart({
      action: 'chart', operation: 'update', target: 'plain_update',
      series: [{ name: 'A', values: [2], axis: 'secondary' }],
    }, ast, parser)).toThrow(/not "combo"/);
    expect(hash(parser.rezip())).toBe(before);
  });
});
