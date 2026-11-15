/**
 * 0.5.3: charts must look right in Word as well as LibreOffice.
 *
 *  - the chart title carries <c:overlay val="0"/> (without it Word draws the title on top of the plot area)
 *  - every <c:majorGridlines> has an explicit line (an empty element is drawn black by Word)
 *
 * Every chart type that has a title or gridlines is covered, on insert, on title update and when gridlines are
 * switched on by a style update.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, executeChart, validateDocxBuffer, type Chart2DSeries } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const SERIES = [
  { name: 'Revenue', values: [10, 20, 30], color: '0066CC' },
  { name: 'Costs', values: [5, 7, 9], color: '10B981' },
];
const COMBO_SERIES = [
  { name: 'Revenue', values: [10, 20, 30], chartType: 'column', color: '0066CC' },
  { name: 'Margin', values: [0.2, 0.3, 0.4], chartType: 'line', axis: 'secondary', color: '10B981' },
];

const TYPES_WITH_GRIDLINES = [
  'column', 'bar', 'line', 'area', 'scatter', 'combo', 'columnStacked', 'barStacked',
  'columnPercentStacked', 'barPercentStacked', 'column3D', 'bar3D', 'line3D',
] as const;
const TYPES_WITHOUT_GRIDLINES = ['pie', 'doughnut', 'pie3D'] as const;

function chartXml(buf: Buffer): string {
  const entry = new AdmZip(buf).getEntry('word/charts/chart1.xml');
  if (!entry) throw new Error('chart1.xml missing');
  return entry.getData().toString('utf8');
}

async function insert(chartType: string, extra: Record<string, unknown> = {}) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const series = (chartType === 'combo' ? COMBO_SERIES : SERIES) as unknown as Chart2DSeries[];
  executeChart(
    { action: 'chart', chartType: chartType as never, id: 'c1', title: 'Quarterly revenue', categories: chartType === 'scatter' ? ['1', '2', '3'] : ['Q1', 'Q2', 'Q3'], series, ...extra },
    ast,
    parser
  );
  return { parser, ast, buf: parser.rezip() };
}

/** The chart-level title block (the one that is a direct child of <c:chart>). */
function chartLevelTitle(xml: string): string {
  const chartStart = xml.indexOf('<c:chart>');
  const start = xml.indexOf('<c:title>', chartStart);
  const end = xml.indexOf('</c:title>', start) + '</c:title>'.length;
  return xml.slice(start, end);
}

describe('chart titles carry overlay 0', () => {
  for (const type of [...TYPES_WITH_GRIDLINES, ...TYPES_WITHOUT_GRIDLINES]) {
    it(`${type}: the chart title is above the plot area, not on it`, async () => {
      const { buf } = await insert(type);
      const title = chartLevelTitle(chartXml(buf));
      expect(title).toContain('<a:t>Quarterly revenue</a:t>');
      expect(title.match(/<c:overlay val="0"\/>/g)).toHaveLength(1);
      // CT_Title child order: tx, layout, overlay, spPr, txPr
      expect(title.indexOf('<c:tx>')).toBeLessThan(title.indexOf('<c:overlay'));
    });
  }

  it('removing the title (explicit empty title) removes the title and its overlay, and stays valid', async () => {
    const { parser, ast } = await insert('column');
    const before = chartXml(parser.rezip());
    executeChart({ action: 'chart', operation: 'update', target: 'c1', title: '' } as never, ast, parser);
    const xml = chartXml(parser.rezip());
    expect(xml).not.toContain('<c:title>');
    // the legend keeps its own overlay; only the title's overlay is gone
    expect((before.match(/<c:overlay/g) ?? []).length - (xml.match(/<c:overlay/g) ?? []).length).toBe(1);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.errors).toEqual([]);
  });

  it('a title text update keeps exactly one overlay and stays valid', async () => {
    const { parser, ast } = await insert('column');
    executeChart({ action: 'chart', operation: 'update', target: 'c1', title: 'A new title' } as never, ast, parser);
    const xml = chartXml(parser.rezip());
    const title = chartLevelTitle(xml);
    expect(title).toContain('<a:t>A new title</a:t>');
    expect(title.match(/<c:overlay val="0"\/>/g)).toHaveLength(1);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.errors).toEqual([]);
  });

  it('adding a title to an untitled chart writes the overlay', async () => {
    const { parser, ast } = await insert('column');
    executeChart({ action: 'chart', operation: 'update', target: 'c1', title: '' } as never, ast, parser);
    executeChart({ action: 'chart', operation: 'update', target: 'c1', title: 'Now titled' } as never, ast, parser);
    const title = chartLevelTitle(chartXml(parser.rezip()));
    expect(title).toContain('<a:t>Now titled</a:t>');
    expect(title.match(/<c:overlay val="0"\/>/g)).toHaveLength(1);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.errors).toEqual([]);
  });

  it('a title with XML characters is escaped', async () => {
    const { buf } = await insert('column', { title: 'R&D <2026> "plan"' });
    const title = chartLevelTitle(chartXml(buf));
    expect(title).toContain('R&amp;D &lt;2026&gt; &quot;plan&quot;');
  });
});

describe('gridlines carry an explicit line', () => {
  for (const type of TYPES_WITH_GRIDLINES) {
    it(`${type}: every gridline element has a light grey line, none is empty`, async () => {
      const { buf } = await insert(type);
      const xml = chartXml(buf);
      expect(xml).not.toContain('<c:majorGridlines/>');
      const blocks = xml.match(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/g) ?? [];
      expect(blocks.length).toBeGreaterThan(0);
      for (const block of blocks) {
        expect(block).toContain('<a:ln w="9525"');
        expect(block).toContain('<a:srgbClr val="D9D9D9"/>');
      }
    });
  }

  for (const type of TYPES_WITHOUT_GRIDLINES) {
    it(`${type}: has no gridlines`, async () => {
      const { buf } = await insert(type);
      expect(chartXml(buf)).not.toContain('majorGridlines');
    });
  }

  it('switching gridlines on with a style update writes the styled line', async () => {
    const { parser, ast } = await insert('column', { style: { axes: { value: { majorGridlines: false } } } });
    expect(chartXml(parser.rezip())).not.toContain('majorGridlines');
    executeChart(
      { action: 'chart', operation: 'update', target: 'c1', style: { axes: { value: { majorGridlines: true } } } } as never,
      ast,
      parser
    );
    const xml = chartXml(parser.rezip());
    const blocks = xml.match(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/g) ?? [];
    expect(blocks).toHaveLength(1);
    expect(blocks[0]).toContain('<a:srgbClr val="D9D9D9"/>');
    expect(xml).not.toContain('<c:majorGridlines/>');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.errors).toEqual([]);
  });

  it('an authored gridline line is not overwritten by a style update', async () => {
    const { parser, ast } = await insert('column');
    const adapter = parser.getZipAdapter();
    const xml = adapter.readText('word/charts/chart1.xml');
    const authored = xml.replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, '<c:majorGridlines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>');
    adapter.writeText('word/charts/chart1.xml', authored);
    executeChart(
      { action: 'chart', operation: 'update', target: 'c1', style: { axes: { value: { majorGridlines: true } } } } as never,
      ast,
      parser
    );
    expect(adapter.readText('word/charts/chart1.xml')).toContain('<a:srgbClr val="FF0000"/>');
  });
});

describe('every generated chart is still valid', () => {
  for (const type of [...TYPES_WITH_GRIDLINES, ...TYPES_WITHOUT_GRIDLINES]) {
    it(`${type}: passes the OOXML validator`, async () => {
      const { buf } = await insert(type);
      const docx = await validateDocxBuffer(buf);
      expect(docx.errors).toEqual([]);
      expect(docx.ok).toBe(true);
    });
  }
});
