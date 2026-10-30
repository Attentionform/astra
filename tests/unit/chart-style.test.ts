import { describe, it, expect } from 'vitest';
import crypto from 'node:crypto';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  DocumentAST,
  OperationSchema,
  ChartType,
  applyChartFloating,
  executeChart,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const CHART_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

const DATA = {
  title: 'Revenue',
  categories: ['Q1', 'Q2'],
  series: [
    { name: 'ARR', values: [10, 20], color: '0066CC' },
    { name: 'Costs', values: [5, 7], color: '10B981' },
  ],
};

function chartXmlOf(buf: Buffer, name = 'chart1.xml'): string {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  if (!entry) throw new Error(`${name} missing from docx`);
  return entry.getData().toString('utf8');
}

function hash(buf: Buffer): string {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

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

async function insertChart(
  id: string,
  opts: Record<string, unknown> = {},
  chartType: ChartType = 'column'
): Promise<{ parser: DocxParser; ast: DocumentAST; id: string; xml: string; buf: Buffer }> {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  executeChart({ action: 'chart', chartType, id, ...DATA, ...opts }, ast, parser);
  const buf = parser.rezip();
  return { parser, ast, id, xml: chartXmlOf(buf), buf };
}

function nodeById(ast: DocumentAST, id: string) {
  return ast.sections[0].nodes.find((n) => n.id === id)!;
}

describe('chart style: legend, numFmt, data labels', () => {
  it('writes legend position, value-axis numFmt and data labels on insert', async () => {
    const { parser, ast, id, xml } = await insertChart('style_insert', {
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
        axes: { value: { numberFormat: '$#,##0' } },
      },
    });

    expect(xml).toContain('<c:legendPos val="t"/>');
    expect(xml).toContain('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>');
    expect(xml).toContain('<c:dLbls><c:numFmt formatCode="$#,##0" sourceLinked="0"/><c:dLblPos val="outEnd"/>');
    expect(xml).toContain('<c:showVal val="1"/>');
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
        axes: { value: { numberFormat: '$#,##0' } },
      },
    });

    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
        axes: { value: { numberFormat: '$#,##0' } },
      },
    });
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('patches style only, keeps data bytes and unparsed styling, and bumps one version', async () => {
    const { parser, ast, id, xml } = await insertChart('style_update');
    const styled = xml
      .replace('<c:legendPos val="r"/>', '<c:legendPos val="b"/>')
      .replace(
        /(<c:ser>\s*<c:idx val="1"\/>[\s\S]*?)(<c:cat>)/,
        '$1<c:trendline><c:trendlineType val="linear"/></c:trendline>$2'
      );
    parser.getZipAdapter().writeText('word/charts/chart1.xml', styled);

    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { legend: { position: 't' }, axes: { value: { numberFormat: '0%' } } },
      },
      ast,
      parser
    );

    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:trendline><c:trendlineType val="linear"/></c:trendline>');
    expect(after).toContain('<c:legendPos val="t"/>');
    expect(after).toContain('<c:numFmt formatCode="0%" sourceLinked="0"/>');
    expect(after).toContain('<c:v>10</c:v>');
    expect(after).toContain('<c:v>20</c:v>');
    expect(nodeById(ast, id).version).toBe(2);
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      style: { legend: { position: 't' }, axes: { value: { numberFormat: '0%' } } },
    });
  });

  it('partial data-label update keeps other authored flags', async () => {
    const { parser, ast, id, xml } = await insertChart('style_pie_labels', {}, 'pie');
    expect(xml).toContain('<c:showPercent val="1"/>');

    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { dataLabels: { showValue: true } },
      },
      ast,
      parser
    );

    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:showPercent val="1"/>');
    expect(after).toContain('<c:showVal val="1"/>');
    expect(after).toContain('<c:showCatName val="0"/>');
  });

  it('removes the legend for position none and leaves data intact', async () => {
    const { parser, ast, id } = await insertChart('style_no_legend');
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { legend: { position: 'none' } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).not.toContain('<c:legend>');
    expect(after).toContain('<c:v>20</c:v>');
  });

  it('patches default-namespace and x-alias parts without converting their dialect', async () => {
    for (const [name, transform] of [
      ['default', toDefaultNamespace],
      ['alias', toXAlias],
    ] as const) {
      const { parser, ast, id, xml } = await insertChart(`style_ns_${name}`);
      parser.getZipAdapter().writeText('word/charts/chart1.xml', transform(xml));
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: id,
          style: { legend: { position: 'tr' }, dataLabels: { show: true, position: 'ctr' } },
        },
        ast,
        parser
      );
      const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
      expect(after).toContain('legendPos val="tr"');
      expect(after).toContain('dLblPos val="ctr"');
      if (name === 'default') {
        expect(after).not.toContain('<c:legend>');
        expect(after).toContain('<legend>');
      } else {
        expect(after).toContain('<x:legend>');
        expect(after).not.toContain('<c:legend>');
      }
    }
  });

  it('applies legend and value numFmt to 3D charts', async () => {
    const { parser, ast, id, xml } = await insertChart(
      'style_3d',
      { style: { legend: { position: 'b' }, axes: { value: { numberFormat: '0.0' } } } },
      'bar3D'
    );
    expect(xml).toContain('<c:legendPos val="b"/>');
    expect(xml).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"/>');
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart3d')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: { legend: { position: 'b' }, axes: { value: { numberFormat: '0.0' } } },
    });
    expect(nodeById(ast, id).version).toBe(1);
  });

  it('accepts the strict style schema and rejects unknown keys and bad enums', () => {
    const base = { action: 'chart', chartType: 'column', series: DATA.series };
    expect(OperationSchema.safeParse({
      ...base,
      style: { legend: { position: 't' }, dataLabels: { show: true }, axes: { value: { numberFormat: '0%' } } },
    }).success).toBe(true);
    expect(OperationSchema.safeParse({ ...base, style: { legend: { position: 'x' } } }).success).toBe(false);
    expect(OperationSchema.safeParse({ ...base, style: { legend: { position: 't', bogus: true } } }).success).toBe(false);
    expect(OperationSchema.safeParse({ ...base, style: { dataLabels: { numberFormat: '' } } }).success).toBe(false);
  });

  it('refuses axes on pie with E_VALIDATION and zero mutation', async () => {
    const { parser, ast, id } = await insertChart('style_pie_refuse', {}, 'pie');
    const before = hash(parser.rezip());
    expect(() => executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { axes: { value: { numberFormat: '0%' } } },
      },
      ast,
      parser
    )).toThrow(/no value axis/);
    expect(hash(parser.rezip())).toBe(before);
  });

  it('refuses invalid style before insert and leaves the package unchanged', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const before = hash(parser.rezip());
    expect(() => executeChart(
      {
        action: 'chart',
        chartType: 'column',
        id: 'bad_style',
        ...DATA,
        style: { dataLabels: { position: 'inside' as never } },
      },
      ast,
      parser
    )).toThrow(/dataLabels.position/);
    expect(hash(parser.rezip())).toBe(before);
    expect(ast.sections[0].nodes.some((n) => n.id === 'bad_style')).toBe(false);
  });
});

describe('chart style: axes and trendlines', () => {
  it('writes axis titles, bounds, gridline removal and a per-series trendline', async () => {
    const { parser, xml, id, ast } = await insertChart('axes_insert', {
      style: {
        axes: {
          category: { title: 'Quarter', tickLabelPosition: 'low' },
          value: {
            title: 'Revenue',
            min: 0,
            max: 100,
            numberFormat: '$#,##0',
            majorGridlines: false,
          },
        },
        trendline: {
          type: 'polynomial',
          series: [2],
          order: 3,
          displayEquation: true,
          displayRSquared: true,
        },
      },
    });
    expect(xml).toContain('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1000"/><a:t>Quarter</a:t>');
    expect(xml).toContain('<a:t>Revenue</a:t>');
    expect(xml).toContain('<c:scaling><c:orientation val="minMax"/><c:max val="100"/><c:min val="0"/></c:scaling>');
    expect(xml).toContain('<c:tickLblPos val="low"/>');
    expect(xml).toContain('<c:numFmt formatCode="$#,##0" sourceLinked="0"/>');
    const valAx = /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(xml)![0];
    expect(valAx).not.toContain('<c:majorGridlines/>');
    const seriesBlocks = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(seriesBlocks[0]).not.toContain('<c:trendline>');
    expect(seriesBlocks[1]).toContain('<c:trendline><c:trendlineType val="polynomial"/><c:order val="3"/><c:dispRSqr val="1"/><c:dispEq val="1"/></c:trendline>');
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      style: {
        axes: {
          category: { title: 'Quarter' },
          value: { title: 'Revenue', min: 0, max: 100, numberFormat: '$#,##0' },
        },
        trendline: { type: 'polynomial', series: [2], order: 3, displayEquation: true, displayRSquared: true },
      },
    });

    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: {
        axes: { category: { title: 'Quarter' }, value: { title: 'Revenue', min: 0, max: 100 } },
        trendline: { type: 'polynomial', series: [2], order: 3 },
      },
    });
  });

  it('numbers series from 1: 1 is the first series, 0 is refused, and the number reads back', async () => {
    const { parser, ast, id } = await insertChart('trend_one_based', {
      style: { trendline: { type: 'linear', series: [1] } },
    });
    const blocks = (): string[] => parser.getZipAdapter().readText('word/charts/chart1.xml').match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(blocks()[0]).toContain('<c:trendline>');
    expect(blocks()[1]).not.toContain('<c:trendline>');
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({ style: { trendline: { type: 'linear', series: [1] } } });

    const before = hash(parser.rezip());
    let failure: unknown;
    try {
      executeChart({ action: 'chart', operation: 'update', target: id, style: { trendline: { type: 'linear', series: [0] } } }, ast, parser);
    } catch (error) {
      failure = error;
    }
    expect(String(failure)).toMatch(/trendline\.series index 0 is outside 1\.\.2/);
    expect((failure as { remedy?: string }).remedy).toMatch(/numbered from 1/);
    expect(hash(parser.rezip())).toBe(before);

    const past = (): unknown => {
      try {
        executeChart({ action: 'chart', operation: 'update', target: id, style: { trendline: { type: 'linear', series: [3] } } }, ast, parser);
      } catch (error) {
        return error;
      }
      return undefined;
    };
    expect(String(past())).toMatch(/index 3 is outside 1\.\.2/);

    const reopened = await new DocxParser(parser.rezip()).getAST();
    const chart = reopened.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({ style: { trendline: { type: 'linear', series: [1] } } });
  });

  it('removes a trendline with type none without touching other series', async () => {
    const { parser, ast, id, xml } = await insertChart('axes_trend_none', {
      style: { trendline: { type: 'linear', series: [1, 2] } },
    });
    expect(xml.match(/<c:trendline>/g)).toHaveLength(2);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { trendline: { type: 'none', series: [2] } },
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const blocks = after.match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(blocks[0]).toContain('<c:trendline>');
    expect(blocks[1]).not.toContain('<c:trendline>');
  });

  it('updates an existing axis title in place and preserves unrelated axis nodes', async () => {
    const { parser, ast, id, xml } = await insertChart('axes_update', {
      style: { axes: { value: { title: 'Old', numberFormat: '0.0' } } },
    });
    const styled = xml.replace('<c:crosses val="autoZero"/>', '<c:crosses val="max"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', styled);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { axes: { value: { title: 'New' } } },
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:t>New</a:t>');
    expect(after).not.toContain('<a:t>Old</a:t>');
    expect(after).toContain('<c:crosses val="max"/>');
    expect(after).toContain('formatCode="0.0"');
  });

  it('styles scatter x as the category axis and y as the value axis', async () => {
    const { xml } = await insertChart(
      'axes_scatter',
      {
        categories: ['0', '1'],
        style: {
          axes: {
            category: { title: 'Load' },
            value: { title: 'Latency', min: 0, max: 100 },
          },
        },
      },
      'scatter'
    );
    expect(xml).toContain('<a:t>Load</a:t>');
    expect(xml).toContain('<a:t>Latency</a:t>');
    const axes = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/g)!;
    expect(axes).toHaveLength(2);
    expect(axes[0]).toContain('<a:t>Load</a:t>');
    expect(axes[1]).toContain('<a:t>Latency</a:t>');
    expect(axes[1]).toContain('<c:max val="100"/>');
  });

  it('accepts axis/trendline schema and rejects invalid combinations', () => {
    const base = { action: 'chart', chartType: 'column', series: DATA.series };
    expect(OperationSchema.safeParse({
      ...base,
      style: {
        axes: { category: { title: 'Q' }, value: { min: 0, max: 10, majorGridlines: false } },
        trendline: { type: 'movingAvg', period: 2 },
      },
    }).success).toBe(true);
    expect(OperationSchema.safeParse({
      ...base,
      style: { trendline: { type: 'polynomial' } },
    }).success).toBe(true); // engine enforces order at apply time
    expect(OperationSchema.safeParse({
      ...base,
      style: { axes: { value: { min: 10, max: 0 } } },
    }).success).toBe(true); // engine enforces the cross-field bound
    expect(OperationSchema.safeParse({
      ...base,
      style: { axes: { value: { tickLabelPosition: 'inside' } } },
    }).success).toBe(false);
  });

  it('refuses polynomial without order, movingAvg without period and out-of-range series before mutation', async () => {
    const { parser, ast, id } = await insertChart('axes_refuse');
    for (const trendline of [
      { type: 'polynomial' as const, series: [1] },
      { type: 'movingAvg' as const, series: [1] },
      { type: 'linear' as const, series: [9] },
      { type: 'linear' as const, series: [0] },
      { type: 'linear' as const, series: [1, 1] },
    ]) {
      const before = hash(parser.rezip());
      expect(() => executeChart(
        { action: 'chart', operation: 'update', target: id, style: { trendline } },
        ast,
        parser
      )).toThrow(/E_VALIDATION/);
      expect(hash(parser.rezip())).toBe(before);
    }
  });

  it('refuses inverted value bounds and 3D trendlines with zero mutation', async () => {
    const { parser, ast, id } = await insertChart('axes_bounds_refuse');
    let before = hash(parser.rezip());
    expect(() => executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { axes: { value: { min: 10, max: 1 } } },
      },
      ast,
      parser
    )).toThrow(/less than/);
    expect(hash(parser.rezip())).toBe(before);

    const three = await insertChart('axes_3d_trend_refuse', {}, 'bar3D');
    before = hash(three.parser.rezip());
    expect(() => executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: three.id,
        style: { trendline: { type: 'linear' } },
      },
      three.ast,
      three.parser
    )).toThrow(/3D chart parts/);
    expect(hash(three.parser.rezip())).toBe(before);
  });
});

describe('chart style: floating anchor', () => {
  function documentXml(parser: DocxParser): string {
    const zip = new AdmZip(parser.rezip());
    return zip.getEntry('word/document.xml')!.getData().toString('utf8');
  }

  it('inserts a floating anchor with alignment and wrap instead of wp:inline', async () => {
    const { parser, xml } = await insertChart('float_insert', {
      style: { floating: { horizontalAlign: 'right', verticalAlign: 'top', behindText: true, wrap: 'square' } },
    });
    const doc = documentXml(parser);
    expect(doc).toContain('<wp:anchor');
    expect(doc).not.toContain('<wp:inline');
    expect(doc).toContain('<wp:align>right</wp:align>');
    expect(doc).toContain('behindDoc="1"');
    expect(doc).toContain('<wp:wrapSquare wrapText="bothSides"/>');
    expect(xml).not.toContain('wp:anchor');

    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chart.metadata.chartData).toMatchObject({
      style: { floating: { horizontalAlign: 'right', verticalAlign: 'top', behindText: true, wrap: 'square' } },
    });
  });

  it('converts an inline drawing to a floating anchor and preserves chart part bytes', async () => {
    const { parser, ast, id, xml } = await insertChart('float_update');
    const beforeChart = parser.getZipAdapter().readText('word/charts/chart1.xml');
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { floating: { horizontalAlign: 'left', verticalAlign: 'bottom', wrap: 'none' } } },
      ast,
      parser
    );
    const doc = documentXml(parser);
    expect(doc).toContain('<wp:anchor');
    expect(doc).not.toContain('<wp:inline');
    expect(doc).toContain('<wp:align>left</wp:align>');
    expect(doc).toContain('<wp:align>bottom</wp:align>');
    expect(doc).toContain('<wp:wrapNone/>');
    expect(parser.getZipAdapter().readText('word/charts/chart1.xml')).toBe(beforeChart);
    expect(nodeById(ast, id).metadata.chartData).toMatchObject({
      style: { floating: { horizontalAlign: 'left', verticalAlign: 'bottom', wrap: 'none' } },
    });
    expect(xml).toContain('<c:legendPos val="r"/>');
  });

  it('repositions an existing anchor without duplicating position or wrap nodes', async () => {
    const { parser, ast, id } = await insertChart('float_reposition', {
      style: { floating: { horizontalAlign: 'center', verticalAlign: 'top', wrap: 'square' } },
    });
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { floating: { horizontalAlign: 'right', verticalAlign: 'middle', wrap: 'square' } } },
      ast,
      parser
    );
    const doc = documentXml(parser);
    expect(doc.match(/<wp:positionH/g)).toHaveLength(1);
    expect(doc.match(/<wp:positionV/g)).toHaveLength(1);
    expect(doc.match(/<wp:wrapSquare/g)).toHaveLength(1);
    expect(doc).toContain('<wp:align>right</wp:align>');
    expect(doc).toContain('<wp:align>center</wp:align>');
  });

  it('accepts floating schema and rejects unknown alignment/wrap', () => {
    const base = { action: 'chart', chartType: 'column', series: DATA.series };
    expect(OperationSchema.safeParse({
      ...base,
      style: { floating: { horizontalAlign: 'right', verticalAlign: 'middle', behindText: false, wrap: 'square' } },
    }).success).toBe(true);
    expect(OperationSchema.safeParse({ ...base, style: { floating: { horizontalAlign: 'inside' } } }).success).toBe(false);
    expect(OperationSchema.safeParse({ ...base, style: { floating: { wrap: 'tight' } } }).success).toBe(false);
  });
});

describe('chart style: schema-correct legend and label patching', () => {
  it('expands a self-closing legend before writing legendPos', async () => {
    const { parser, ast, id, xml } = await insertChart('legend_self_closing');
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      xml.replace(/<c:legend>[\s\S]*?<\/c:legend>/, '<c:legend/>')
    );
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { legend: { position: 't' } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:legend><c:legendPos val="t"/></c:legend>');
    expect(after).not.toContain('<c:legend/><c:legendPos');
  });

  it('inserts overlay before spPr in an existing legend', async () => {
    const { parser, ast, id, xml } = await insertChart('legend_overlay_order');
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      xml.replace(
        /<c:legend>[\s\S]*?<\/c:legend>/,
        '<c:legend><c:layout/><c:spPr><a:noFill/></c:spPr></c:legend>'
      )
    );
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { legend: { overlay: true } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:legend><c:layout/><c:overlay val="1"/><c:spPr><a:noFill/></c:spPr></c:legend>');
  });

  it('inserts a new legend after plotArea and before dispBlanksAs', async () => {
    const { parser, ast, id, xml } = await insertChart('legend_after_plot_area');
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      xml.replace(/<c:legend>[\s\S]*?<\/c:legend>/, '<c:dispBlanksAs val="gap"/>')
    );
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { legend: { position: 'b' } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const plotClose = after.indexOf('</c:plotArea>');
    const legendAt = after.indexOf('<c:legend>');
    const blanksAt = after.indexOf('<c:dispBlanksAs');
    expect(plotClose).toBeGreaterThan(-1);
    expect(legendAt).toBeGreaterThan(plotClose);
    expect(blanksAt).toBeGreaterThan(legendAt);
    expect(after).toContain('<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>');
  });

  it('clears c:delete when data labels are turned on', async () => {
    const { parser, ast, id, xml } = await insertChart('labels_delete_clear', {}, 'pie');
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      xml.replace('<c:dLbls>', '<c:dLbls><c:delete val="1"/>')
    );
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { dataLabels: { show: true } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:delete val="0"/>');
    expect(after).toContain('<c:showVal val="1"/>');
  });

  it('treats an empty dataLabels object as a no-op', async () => {
    const { parser, ast, id, xml } = await insertChart('labels_empty_noop');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', xml);
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { dataLabels: {} } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toBe(xml);
    expect(after).not.toContain('<c:dLbls>');
  });

  it('sets c:delete when data labels are turned off', async () => {
    const { parser, ast, id } = await insertChart('labels_delete_set', {}, 'pie');
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { dataLabels: { show: false } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:delete val="1"/>');
    expect(after).not.toContain('<c:showVal val="1"/>');
    expect(after).toContain('<c:showPercent val="0"/>');
  });
});

describe('chart style: axis child order', () => {
  it('keeps tickLblPos before spPr and crossAx on a minimal category axis', async () => {
    const { parser, ast, id, xml } = await insertChart('axis_order_category');
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      xml.replace(
        /<c:catAx>[\s\S]*?<\/c:catAx>/,
        '<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:spPr><a:noFill/></c:spPr><c:crossAx val="1002"/></c:catAx>'
      )
    );
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { axes: { category: { tickLabelPosition: 'high' } } } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const catAx = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(after)![0];
    const tickAt = catAx.indexOf('<c:tickLblPos');
    const spPrAt = catAx.indexOf('<c:spPr>');
    const crossAt = catAx.indexOf('<c:crossAx');
    expect(tickAt).toBeGreaterThan(-1);
    expect(spPrAt).toBeGreaterThan(tickAt);
    expect(crossAt).toBeGreaterThan(spPrAt);
  });

  it('keeps title and numFmt before spPr and crossAx on a minimal value axis', async () => {
    const { parser, ast, id, xml } = await insertChart('axis_order_value');
    parser.getZipAdapter().writeText(
      'word/charts/chart1.xml',
      xml.replace(
        /<c:valAx>[\s\S]*?<\/c:valAx>/,
        '<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:spPr><a:noFill/></c:spPr><c:crossAx val="1001"/></c:valAx>'
      )
    );
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        style: { axes: { value: { title: 'Revenue', numberFormat: '0.0' } } },
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const valAx = /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(after)![0];
    const titleAt = valAx.indexOf('<c:title>');
    const numFmtAt = valAx.indexOf('<c:numFmt');
    const spPrAt = valAx.indexOf('<c:spPr>');
    const crossAt = valAx.indexOf('<c:crossAx');
    expect(titleAt).toBeGreaterThan(-1);
    expect(numFmtAt).toBeGreaterThan(titleAt);
    expect(spPrAt).toBeGreaterThan(numFmtAt);
    expect(crossAt).toBeGreaterThan(spPrAt);
  });
});

describe('chart style: legacy defaults are not fabricated', () => {
  it('reads no style from a plain column insert after reopen', async () => {
    // Authoring exactly the generated defaults is indistinguishable from the
    // legacy baseline; plain-insert consistency wins over that duplicate case.
    const { parser } = await insertChart('plain_style_baseline');
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect((chart.metadata.chartData as Record<string, unknown>).style).toBeUndefined();
  });

  it('still reads authored styling from a styled insert after reopen', async () => {
    const { parser } = await insertChart('styled_style_baseline', {
      style: { legend: { position: 't' }, axes: { value: { numberFormat: '$#,##0' } } },
    });
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chartNode = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chartNode.metadata.chartData).toMatchObject({
      style: { legend: { position: 't' }, axes: { value: { numberFormat: '$#,##0' } } },
    });
  });

  it('reads an authored majorGridlines removal instead of treating it as the baseline', async () => {
    const { parser } = await insertChart('styled_gridlines_off', {
      style: { axes: { value: { majorGridlines: false } } },
    });
    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const chartNode = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(chartNode.metadata.chartData).toMatchObject({
      style: { axes: { value: { majorGridlines: false } } },
    });
  });

  it('preserves style in metadata when updating title or series with omitted style', async () => {
    const { parser, ast, id } = await insertChart('styled_preserve_metadata', {
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
      },
    });

    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        title: 'Preserved Title Update',
      },
      ast,
      parser
    );

    let chartNode = nodeById(ast, id);
    expect(chartNode.metadata.chartData).toMatchObject({
      title: 'Preserved Title Update',
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
      },
    });

    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [12, 22] },
          { name: 'Costs', values: [6, 8] },
        ],
      },
      ast,
      parser
    );

    chartNode = nodeById(ast, id);
    expect(chartNode.metadata.chartData).toMatchObject({
      title: 'Preserved Title Update',
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
      },
      series: [
        { name: 'ARR', values: [12, 22] },
        { name: 'Costs', values: [6, 8] },
      ],
    });

    const reopened = new DocxParser(parser.rezip());
    const reopenedAst = await reopened.getAST();
    const reopenedChart = reopenedAst.sections[0].nodes.find((n) => n.type === 'chart')!;
    expect(reopenedChart.metadata.chartData).toMatchObject({
      style: {
        legend: { position: 't', overlay: false },
        dataLabels: { show: true, position: 'outEnd', numberFormat: '$#,##0' },
      },
    });
  });
});

describe('chart style: floating anchor targeting and partial updates', () => {
  function parseElement(xml: string): Element {
    return new DOMParser().parseFromString(xml, 'text/xml').documentElement as Element;
  }

  function drawing(chartRId?: string): string {
    const graphic = chartRId === undefined
      ? '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic>'
      : `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${chartRId}"/></a:graphicData></a:graphic>`;
    return `<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="d"/><wp:wrapNone/>${graphic}</wp:inline></w:drawing>`;
  }

  function paragraph(...drawings: string[]): Element {
    return parseElement(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${drawings.map((d) => `<w:r>${d}</w:r>`).join('')}</w:p>`);
  }

  it('targets the drawing that references the chart, not the first drawing', () => {
    const node = paragraph(drawing(), drawing('rId5'));
    applyChartFloating(node, { horizontalAlign: 'left' }, 'word/charts/chart1.xml', 'rId5');
    const drawings = node.getElementsByTagName('w:drawing');
    expect(drawings.length).toBe(2);
    expect(node.getElementsByTagName('wp:inline').length).toBe(1);
    expect(node.getElementsByTagName('wp:anchor').length).toBe(1);
    expect(node.getElementsByTagName('wp:anchor')[0].textContent).toContain('left');
  });

  it('refuses with E_VALIDATION when no drawing references the chart part', () => {
    const node = paragraph(drawing());
    let caught: unknown;
    try {
      applyChartFloating(node, { horizontalAlign: 'left' }, 'word/charts/chart1.xml', 'rId5');
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(Error);
    expect((caught as Error).message).toMatch(/no drawing that references the chart part/);
    expect((caught as { code?: string }).code).toBe('E_VALIDATION');
    expect((caught as { remedy?: string }).remedy).toContain('delete and re-insert the chart');
  });

  it('preserves unspecified anchor placement on a partial floating update', async () => {
    const { parser, ast, id } = await insertChart('float_partial', {
      style: { floating: { horizontalAlign: 'right', verticalAlign: 'bottom', wrap: 'none' } },
    });
    executeChart(
      { action: 'chart', operation: 'update', target: id, style: { floating: { behindText: true } } },
      ast,
      parser
    );
    const zip = new AdmZip(parser.rezip());
    const doc = zip.getEntry('word/document.xml')!.getData().toString('utf8');
    expect(doc).toContain('<wp:align>right</wp:align>');
    expect(doc).toContain('<wp:align>bottom</wp:align>');
    expect(doc).toContain('<wp:wrapNone/>');
    expect(doc).toContain('behindDoc="1"');
    expect(doc).not.toContain('<wp:align>center</wp:align>');
    expect(doc).not.toContain('<wp:wrapSquare');
  });

  it('preserves cNvGraphicFramePr frame locks and extLst in schema order when converting inline chart to floating', () => {
    const inlineXml = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `
      + `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" `
      + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `
      + `xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" `
      + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
      + `<w:r><w:drawing><wp:inline>`
      + `<wp:extent cx="5486400" cy="3200400"/>`
      + `<wp:effectExtent l="0" t="0" r="0" b="0"/>`
      + `<wp:docPr id="42" name="Chart 42"/>`
      + `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
      + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId8"/></a:graphicData></a:graphic>`
      + `<wp:extLst><ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"/></wp:extLst>`
      + `</wp:inline></w:drawing></w:r></w:p>`;
    const node = parseElement(inlineXml);

    applyChartFloating(
      node,
      { horizontalAlign: 'center', verticalAlign: 'top', wrap: 'square' },
      'word/charts/chart1.xml',
      'rId8'
    );

    expect(node.getElementsByTagName('wp:inline').length).toBe(0);
    const anchors = node.getElementsByTagName('wp:anchor');
    expect(anchors.length).toBe(1);
    const anchor = anchors[0];

    const cNvPr = anchor.getElementsByTagName('wp:cNvGraphicFramePr');
    expect(cNvPr.length).toBe(1);
    const locks = cNvPr[0].getElementsByTagName('a:graphicFrameLocks');
    expect(locks.length).toBe(1);
    expect(locks[0].getAttribute('noChangeAspect')).toBe('1');

    const extLst = anchor.getElementsByTagName('wp:extLst');
    expect(extLst.length).toBe(1);
    expect(extLst[0].getElementsByTagName('ext').length).toBe(1);

    const childTagNames: string[] = [];
    for (let i = 0; i < anchor.childNodes.length; i++) {
      const child = anchor.childNodes[i] as Element;
      if (child.nodeType === 1) {
        childTagNames.push(child.localName ?? child.nodeName.split(':').pop()!);
      }
    }
    expect(childTagNames).toEqual([
      'simplePos',
      'positionH',
      'positionV',
      'extent',
      'effectExtent',
      'wrapSquare',
      'docPr',
      'cNvGraphicFramePr',
      'graphic',
      'extLst',
    ]);
  });
});
