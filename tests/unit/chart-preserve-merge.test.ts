import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, executeChart, toError, validateDocxBuffer } from '@attentionform/astra-engine';

/**
 * Preserve-and-merge chart update pins: a data-only update patches the data
 * nodes inside the existing chart part instead of regenerating it, so every
 * node the read path does not model (axes config, ticks, numFmt, legend
 * placement, per-point formats and fills, trendlines, plot fills) survives
 * byte-identical. Unpatchable shapes refuse with E_VALIDATION + remedy and
 * mutate nothing.
 */

const STYLE_MARKERS = [
  '<c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></c:spPr>',
  '<c:numFmt formatCode="$#,##0" sourceLinked="0"/>',
  '<c:legendPos val="t"/>',
  '<a:t>Quarter</a:t>',
  'formatCode="$#,##0"',
  '<a:srgbClr val="FF00FF"/>',
  '<c:trendline><c:trendlineType val="linear"/></c:trendline>',
];

/** Injects schema-ordered nodes the chart read path never parses. */
function withStyling(xml: string): string {
  let out = xml;
  out = out.replace(
    '    </c:plotArea>',
    '      <c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></c:spPr>\n    </c:plotArea>'
  );
  out = out.replace(
    '<c:numFmt formatCode="General" sourceLinked="1"/>',
    '<c:numFmt formatCode="$#,##0" sourceLinked="0"/>'
  );
  out = out.replace('<c:legendPos val="r"/>', '<c:legendPos val="t"/>');
  out = out.replace(
    '        <c:axPos val="b"/>\n        <c:majorTickMark val="out"/>',
    '        <c:axPos val="b"/>\n        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1" sz="1000"/><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title>\n        <c:majorTickMark val="out"/>'
  );
  out = out.replace('<c:pt idx="1"><c:v>20</c:v></c:pt>', '<c:pt idx="1" formatCode="$#,##0"><c:v>20</c:v></c:pt>');
  out = out.replace(
    /(<c:ser>\s*<c:idx val="0"\/>[\s\S]*?)(<c:cat>)/,
    '$1<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></c:spPr></c:dPt>$2'
  );
  out = out.replace(
    /(<c:ser>\s*<c:idx val="1"\/>[\s\S]*?)(<c:cat>)/,
    '$1<c:trendline><c:trendlineType val="linear"/></c:trendline>$2'
  );
  return out;
}

function expectStyling(xml: string): void {
  for (const marker of STYLE_MARKERS) expect(xml).toContain(marker);
}

async function styledColumnFixture(prefix: string) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const id = `${prefix}_chart`;
  executeChart(
    {
      action: 'chart',
      chartType: 'column',
      title: 'Styled Title',
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'ARR', values: [10, 20], color: '0066CC' },
        { name: 'Costs', values: [5, 7], color: '10B981' },
      ],
      id,
    },
    ast,
    parser
  );
  const before = withStyling(parser.getZipAdapter().readText('word/charts/chart1.xml'));
  parser.getZipAdapter().writeText('word/charts/chart1.xml', before);
  return { parser, ast, before, id };
}

describe('chart update preserve-and-merge', () => {
  it('data-only update changes one value span and keeps styling bytes in place', async () => {
    const { parser, ast, before, id } = await styledColumnFixture('merge_a');
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [99, 20] },
          { name: 'Costs', values: [5, 7] },
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
    expectStyling(after);
    expect(ast.sections[0].nodes.find((n) => n.id === id)!.version).toBe(2);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
    const ooxml = await parser.validateOOXML();
    expect(ooxml.ok).toBe(true);
    expect(ooxml.errors).toEqual([]);
  });

  it('title, categories, series name and color patch only their own nodes', async () => {
    const { parser, ast, id } = await styledColumnFixture('merge_b');
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        title: 'Restyled Title',
        categories: ['H1', 'H2'],
        series: [
          { name: 'ARR', values: [99, 20], color: 'FF0000' },
          { name: 'Costs', values: [5, 7] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:t>Restyled Title</a:t>');
    expect(after).not.toContain('<a:t>Styled Title</a:t>');
    expect(after).toContain('<c:v>H1</c:v>');
    expect(after).not.toContain('<c:v>Q1</c:v>');
    expect(after).toContain('<a:srgbClr val="FF0000"/>');
    expect(after).not.toContain('<a:srgbClr val="0066CC"/>');
    expect(after).toContain('<a:srgbClr val="10B981"/>');
    expect(after.match(/<c:title>/g)).toHaveLength(2);
    expect(after).toContain('<c:externalData r:id="rId1">');
    expect(after).toContain('<c:overlay val="0"/>');
    expectStyling(after);
  });

  it('refuses a series with no cached values, leaving the part and version untouched', async () => {
    const { parser, ast, id } = await styledColumnFixture('merge_c');
    const stripped = parser
      .getZipAdapter()
      .readText('word/charts/chart1.xml')
      .replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, '');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', stripped);
    const version = ast.sections[0].nodes.find((n) => n.id === id)!.version;
    let refusal: (Error & { code?: string; retryable?: boolean; remedy?: string }) | null = null;
    try {
      executeChart(
        {
          action: 'chart',
          operation: 'update',
          target: id,
          series: [
            { name: 'ARR', values: [9, 20] },
            { name: 'Costs', values: [5, 7] },
          ],
        },
        ast,
        parser
      );
      expect.unreachable('update without a cached value block must refuse');
    } catch (err) {
      refusal = err as Error & { code?: string; retryable?: boolean; remedy?: string };
    }
    expect(refusal!.code).toBe('E_VALIDATION');
    expect(refusal!.retryable).toBe(false);
    expect(refusal!.message).toContain('no cached numeric values');
    expect(refusal!.remedy).toContain('inspect_document');
    expect(refusal!.remedy).toContain("delete it and re-insert it with action:'chart'");
    // The authored remedy is remedy text, not a labelled sentence; wrapping it
    // must not produce `Remedy: Remedy:`.
    expect(refusal!.remedy).not.toMatch(/^Remedy:/);
    const wrapped = toError('E_VALIDATION', refusal!.message, refusal!.remedy);
    expect(wrapped.message.match(/Remedy:/g)).toHaveLength(1);
    expect(wrapped.remedy).toBe(refusal!.remedy);
    expect(parser.getZipAdapter().readText('word/charts/chart1.xml')).toBe(stripped);
    expect(ast.sections[0].nodes.find((n) => n.id === id)!.version).toBe(version);
  });

  it('refuses string labels over a numeric category cache, zero mutation', async () => {
    const { parser, ast, id } = await styledColumnFixture('merge_d');
    const numeric = parser
      .getZipAdapter()
      .readText('word/charts/chart1.xml')
      .replace('<c:cat><c:strLit>', '<c:cat><c:numLit>')
      .replace('</c:strLit></c:cat>', '</c:numLit></c:cat>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', numeric);
    let refusal: (Error & { code?: string; remedy?: string }) | null = null;
    try {
      executeChart(
        { action: 'chart', operation: 'update', target: id, categories: ['H1', 'H2'] },
        ast,
        parser
      );
      expect.unreachable('numeric category cache must refuse string labels');
    } catch (err) {
      refusal = err as Error & { code?: string; remedy?: string };
    }
    expect(refusal!.code).toBe('E_VALIDATION');
    expect(refusal!.message).toContain('numeric categories');
    expect(refusal!.remedy).toContain('numeric labels');
    expect(parser.getZipAdapter().readText('word/charts/chart1.xml')).toBe(numeric);
  });

  it('inserts a chart title when the part carries none, leaving the axis title intact', async () => {
    const { parser, ast, id, before } = await styledColumnFixture('merge_f');
    const titleStart = before.indexOf('    <c:title>');
    const titleEnd = before.indexOf('</c:title>', titleStart) + '</c:title>'.length + 1;
    const withoutChartTitle = before.slice(0, titleStart) + before.slice(titleEnd);
    expect(withoutChartTitle).toContain('<a:t>Quarter</a:t>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', withoutChartTitle);
    executeChart(
      { action: 'chart', operation: 'update', target: id, title: 'Added Title' },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:t>Added Title</a:t>');
    expect(after).toContain('<a:t>Quarter</a:t>');
    expect(after.match(/<c:title>/g)).toHaveLength(2);
    const start = after.indexOf('<c:title>');
    const end = after.indexOf('</c:title>', start) + '</c:title>'.length;
    expect(after.slice(end, end + 5)).toBe('\n    ');
    expect(after.slice(0, start) + after.slice(end + 5)).toBe(withoutChartTitle);
  });

  it('series-count changes keep the structural regeneration path', async () => {
    const { parser, ast, id } = await styledColumnFixture('merge_e');
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [99, 20] },
          { name: 'Costs', values: [5, 7] },
          { name: 'Third', values: [1, 2], color: 'F59E0B' },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:v>Third</c:v>');
    expect(after).toContain('<a:srgbClr val="F59E0B"/>');
    expect(after.match(/<c:ser>/g)).toHaveLength(3);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });
});

describe('chart update preserve-and-merge: 3D and pie', () => {
  it('patches 3D bar shape and inserts camera keys in schema order', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart(
      { action: 'chart', chartType: 'bar3D', shape: 'box', title: '3D', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'merge_3d' },
      ast,
      parser
    );
    executeChart({ action: 'chart', operation: 'update', target: 'merge_3d', shape: 'cone', view3D: { depthPercent: 140 } }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:shape val="cone"/>');
    expect(after).not.toContain('<c:shape val="box"/>');
    expect(after).toContain('<c:rotX val="30"/>');
    expect(after).toContain('<c:perspective val="30"/>');
    expect(after.indexOf('<c:depthPercent val="140"/>')).toBeLessThan(after.indexOf('<c:rAngAx val="1"/>'));
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('keeps per-point pie slice fills byte-identical when values change', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart(
      { action: 'chart', chartType: 'pie', title: 'Pie', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'merge_pie' },
      ast,
      parser
    );
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const slicesBefore = [...before.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)].map((m) => m[0]);
    expect(slicesBefore.length).toBe(2);
    executeChart({ action: 'chart', operation: 'update', target: 'merge_pie', series: [{ name: 'S', values: [3, 2] }] }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const slicesAfter = [...after.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)].map((m) => m[0]);
    expect(slicesAfter).toEqual(slicesBefore);
    expect(after).toContain('<c:v>3</c:v>');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
  });
});




describe('chart update preserve-and-merge: rework pins', () => {
  it('leaves unparsed fills byte-identical on a values-only update', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_a');
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const foreign = before
      .replace('<a:srgbClr val="0066CC"/>', '<a:schemeClr val="accent1"/>')
      .replace(
        '<a:solidFill>\n              <a:srgbClr val="10B981"/>\n            </a:solidFill>',
        '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="111111"/></a:gs><a:gs pos="100000"><a:srgbClr val="222222"/></a:gs></a:gsLst></a:gradFill>' +
          '<a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>'
      );
    expect(foreign).toContain('<a:schemeClr val="accent1"/>');
    expect(foreign).toContain('<a:gradFill>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', foreign);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [99, 20] },
          { name: 'Costs', values: [5, 7] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:v>99</c:v>');
    expect(after).toContain('<a:schemeClr val="accent1"/>');
    expect(after).toContain('<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="111111"/></a:gs><a:gs pos="100000"><a:srgbClr val="222222"/></a:gs></a:gsLst></a:gradFill>');
    expect(after).toContain('<a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
  });

  it('patches the chart-level 3D shape and never a per-series shape', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart(
      { action: 'chart', chartType: 'bar3D', shape: 'box', title: '3D', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'rework_b' },
      ast,
      parser
    );
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const seeded = before.replace('</c:val>\n        </c:ser>', '</c:val>\n          <c:shape val="cylinder"/>\n        </c:ser>');
    expect(seeded).toContain('<c:shape val="cylinder"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', seeded);
    executeChart({ action: 'chart', operation: 'update', target: 'rework_b', shape: 'cone' }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:shape val="cylinder"/>');
    expect(after).toContain('<c:shape val="cone"/>');
    expect(after.indexOf('val="cylinder"')).toBeLessThan(after.indexOf('val="cone"'));
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('never restyles a nested per-point fill when an explicit color is provided', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_c');
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const noSeriesFill = before.replace(
      '<c:spPr>\n            <a:solidFill>\n              <a:srgbClr val="0066CC"/>\n            </a:solidFill>\n          </c:spPr>\n          ',
      ''
    );
    expect(noSeriesFill).not.toContain('<a:srgbClr val="0066CC"/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', noSeriesFill);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [10, 20], color: 'FF0000' },
          { name: 'Costs', values: [5, 7] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:srgbClr val="FF00FF"/>');
    expect(after).toContain('<a:srgbClr val="FF0000"/>');
    expect(after.indexOf('val="FF0000"')).toBeLessThan(after.indexOf('val="FF00FF"'));
    expect([...after.matchAll(/val="FF0000"/g)]).toHaveLength(1);
  });

  it('applies a requested title that equals the read-path default or an axis title', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_d');
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const chartTitleStart = before.indexOf('    <c:title>');
    const chartTitleEnd = before.indexOf('</c:title>', chartTitleStart) + '</c:title>'.length + 1;
    const noChartTitle = before.slice(0, chartTitleStart) + before.slice(chartTitleEnd);
    parser.getZipAdapter().writeText('word/charts/chart1.xml', noChartTitle);
    executeChart({ action: 'chart', operation: 'update', target: id, title: 'Quarter' }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after.match(/<c:title>/g)).toHaveLength(2);
    expect(after.match(/<a:t>Quarter<\/a:t>/g)).toHaveLength(2);
  });

  it('inserts a title when the requested text equals the read default', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_e');
    let xml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    for (let i = 0; i < 2; i++) {
      const start = xml.indexOf('<c:title>');
      const end = xml.indexOf('</c:title>', start) + '</c:title>'.length;
      xml = xml.slice(0, start) + xml.slice(end);
    }
    expect(xml).not.toContain('<c:title>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', xml);
    executeChart({ action: 'chart', operation: 'update', target: id, title: 'Chart Analysis' }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:t>Chart Analysis</a:t>');
  });

  it('replaces a linked (strRef) title without duplicating its c:tx', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_h');
    const xml = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const start = xml.indexOf('<c:title>');
    const end = xml.indexOf('</c:title>', start) + '</c:title>'.length;
    const linked =
      '<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Linked</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val="0"/></c:title>';
    const seeded = xml.slice(0, start) + linked + xml.slice(end);
    const txCount = (seeded.match(/<c:tx>/g) || []).length;
    parser.getZipAdapter().writeText('word/charts/chart1.xml', seeded);
    executeChart({ action: 'chart', operation: 'update', target: id, title: 'New Title' }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:t>New Title</a:t>');
    expect(after).not.toContain('<c:strRef>');
    expect(after.match(/<c:tx>/g) || []).toHaveLength(txCount);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('replaces only the fill when an explicit color lands on a gradient series', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_i');
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const gradient = before.replace(
      '<a:solidFill>\n              <a:srgbClr val="10B981"/>\n            </a:solidFill>',
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs><a:gs pos="100000"><a:srgbClr val="ED7D31"/></a:gs></a:gsLst></a:gradFill>' +
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>'
    );
    expect(gradient).not.toBe(before);
    parser.getZipAdapter().writeText('word/charts/chart1.xml', gradient);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [10, 20] },
          { name: 'Costs', values: [5, 7], color: 'FF0000' },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:srgbClr val="FF0000"/>');
    expect(after).not.toContain('<a:gradFill>');
    expect(after).toContain('<a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>');
    expect(after).toContain('<a:srgbClr val="0066CC"/>');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('fills a self-closed series spPr instead of dropping an explicit color', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_l');
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const seeded = before.replace(/<c:spPr>[\s\S]*?<\/c:spPr>/, '<c:spPr/>');
    expect(seeded).toContain('<c:spPr/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', seeded);
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [10, 20], color: 'FF0000' },
          { name: 'Costs', values: [5, 7] },
        ],
      },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>');
    expect(after).not.toContain('<c:spPr/>');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('fills a self-closed a:ln instead of dropping a line stroke color', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart(
      { action: 'chart', chartType: 'line', title: 'Line', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'rework_m' },
      ast,
      parser
    );
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const seeded = before.replace(/<a:ln w="25400">[\s\S]*?<\/a:ln>/, '<a:ln/>');
    expect(seeded).toContain('<a:ln/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', seeded);
    executeChart(
      { action: 'chart', operation: 'update', target: 'rework_m', series: [{ name: 'S', values: [3, 2], color: 'FF0000' }] },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>');
    expect(after).not.toContain('0066CC');
    expect(after.match(/val="FF0000"/g) || []).toHaveLength(2);
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('writes val onto valueless self-closed shape and ptCount elements', async () => {
    {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      executeChart(
        { action: 'chart', chartType: 'bar3D', shape: 'box', title: '3D', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'rework_n' },
        ast,
        parser
      );
      const raw = parser.getZipAdapter().readText('word/charts/chart1.xml');
      parser.getZipAdapter().writeText('word/charts/chart1.xml', raw.replace('<c:shape val="box"/>', '<c:shape/>'));
      executeChart({ action: 'chart', operation: 'update', target: 'rework_n', shape: 'pyramid' }, ast, parser);
      const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
      expect(after).toContain('<c:shape val="pyramid"/>');
      expect(after).not.toContain('<c:shape/');
      const docx = await validateDocxBuffer(parser.rezip());
      expect(docx.ok).toBe(true);
    }
    {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      executeChart(
        { action: 'chart', chartType: 'column', title: 'T', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'rework_o' },
        ast,
        parser
      );
      const raw = parser.getZipAdapter().readText('word/charts/chart1.xml');
      parser.getZipAdapter().writeText('word/charts/chart1.xml', raw.replace('</c:formatCode><c:ptCount val="2"/>', '</c:formatCode><c:ptCount/>'));
      executeChart({ action: 'chart', operation: 'update', target: 'rework_o', series: [{ name: 'S', values: [3, 2] }] }, ast, parser);
      const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
      expect(after).toContain('<c:ptCount val="2"/>');
      const docx = await validateDocxBuffer(parser.rezip());
      expect(docx.ok).toBe(true);
    }
  });

  it('pins categories-only growth workbook row behaviour (HEAD-identical)', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_p');
    executeChart({ action: 'chart', operation: 'update', target: id, categories: ['Q1', 'Q2', 'Q3'] }, ast, parser);
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect([...after.matchAll(/<c:ptCount val="(\d+)"\/>/g)].map((m) => m[1])).toEqual(['3', '2', '3', '2']);
    const parts = new AdmZip(parser.rezip());
    const wb = new AdmZip(parts.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData());
    const sheet = wb.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8');
    expect([...sheet.matchAll(/<row r="(\d+)"/g)].map((m) => m[1])).toEqual(['1', '2', '3', '4']);
    expect(sheet).toContain('<c r="A4" t="inlineStr"><is><t>Q3</t></is></c>');
    expect(sheet).toContain('<c r="B4"><v>0</v></c>');
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
  });

  it('applies an explicit line-series color to the stroke and marker, not a dead fill', async () => {
    for (const [chartType, id] of [['line', 'rework_j'], ['line3D', 'rework_j3']] as const) {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      executeChart(
        { action: 'chart', chartType, title: 'Line', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id },
        ast,
        parser
      );
      executeChart(
        { action: 'chart', operation: 'update', target: id, series: [{ name: 'S', values: [3, 2], color: 'FF0000' }] },
        ast,
        parser
      );
      const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
      expect(after).toContain('<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>');
      expect(after).not.toContain('0066CC');
      expect(after).toContain('<c:v>3</c:v>');
      expect(after.match(/val="FF0000"/g) || []).toHaveLength(chartType === 'line' ? 2 : 1);
      const docx = await validateDocxBuffer(parser.rezip());
      expect(docx.ok).toBe(true);
      expect(docx.errors).toEqual([]);
    }
  });

  it('fills a self-closed view3D block instead of dropping the camera update', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart(
      { action: 'chart', chartType: 'bar3D', shape: 'box', title: '3D', categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2], color: '0066CC' }], id: 'rework_k' },
      ast,
      parser
    );
    const before = parser.getZipAdapter().readText('word/charts/chart1.xml');
    const stripped = before.replace(/<c:view3D>[\s\S]*?<\/c:view3D>/, '<c:view3D/>');
    expect(stripped).toContain('<c:view3D/>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', stripped);
    executeChart(
      { action: 'chart', operation: 'update', target: 'rework_k', view3D: { depthPercent: 140, perspective: 20 } },
      ast,
      parser
    );
    const after = parser.getZipAdapter().readText('word/charts/chart1.xml');
    expect(after).toContain('<c:depthPercent val="140"/>');
    expect(after).toContain('<c:perspective val="20"/>');
    expect(after.indexOf('<c:depthPercent val="140"/>')).toBeLessThan(after.indexOf('<c:perspective val="20"/>'));
    const docx = await validateDocxBuffer(parser.rezip());
    expect(docx.ok).toBe(true);
    expect(docx.errors).toEqual([]);
  });

  it('keeps workbook rows aligned when a series-only update grows the values', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_f');
    executeChart(
      {
        action: 'chart',
        operation: 'update',
        target: id,
        series: [
          { name: 'ARR', values: [99, 20, 42] },
          { name: 'Costs', values: [5, 7, 8] },
        ],
      },
      ast,
      parser
    );
    const parts = new AdmZip(parser.rezip());
    const wb = new AdmZip(parts.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData());
    const sheet = wb.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8');
    expect(sheet).toContain('<c r="B4"><v>42</v></c>');
    expect(sheet).toContain('<c r="C4"><v>8</v></c>');
    expect(sheet).toContain('<row r="4">');
  });

  it('carries numeric category labels into the regenerated workbook', async () => {
    const { parser, ast, id } = await styledColumnFixture('rework_g');
    const numeric = parser
      .getZipAdapter()
      .readText('word/charts/chart1.xml')
      .split('<c:cat><c:strLit>').join('<c:cat><c:numLit>')
      .split('</c:strLit></c:cat>').join('</c:numLit></c:cat>')
      .split('<c:v>Q1</c:v>').join('<c:v>2024</c:v>')
      .split('<c:v>Q2</c:v>').join('<c:v>2025</c:v>');
    parser.getZipAdapter().writeText('word/charts/chart1.xml', numeric);
    executeChart({ action: 'chart', operation: 'update', target: id, title: 'Numeric cats' }, ast, parser);
    const parts = new AdmZip(parser.rezip());
    const wb = new AdmZip(parts.getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData());
    const sheet = wb.getEntry('xl/worksheets/sheet1.xml')!.getData().toString('utf8');
    expect(sheet).toContain('2024');
    expect(sheet).toContain('2025');
    expect(parser.getZipAdapter().readText('word/charts/chart1.xml')).toContain('<c:numLit>');
  });
});
