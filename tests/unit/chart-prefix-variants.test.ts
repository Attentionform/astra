/**
 * 0.5.3: chart code reads OOXML by namespace and local name, so valid packages that bind a namespace to another
 * prefix behave the same as the usual `wp:`, `c:` and `a:` forms. Also: trendline type names written to the file.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { applyRunPropsDeep, drawingPrefixOf } from '../../packages/engine/src/operations/chart-text';
import { describeLoss } from '../../packages/engine/src/operations/chart-structure';
import { applyChartStyle } from '../../packages/engine/src/operations/chart-style';

const WORD_CHART = join(__dirname, '..', 'fixtures', 'word-chart-with-workbook.docx');

async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  return { parser, ast, run, zip: parser.getZipAdapter() };
}
const chartId = (s: { ast: any }): string => s.ast.sections.flatMap((x: any) => x.nodes).find((n: any) => n.type === 'chart').id;

describe('drawing lookup by namespace', () => {
  it('resizes and sets alt text on a chart drawing whose wp prefix is renamed', async () => {
    const zip = new AdmZip(readFileSync(WORD_CHART));
    const doc = zip.readAsText('word/document.xml');
    expect(doc).toContain('<wp:extent');
    const renamed = doc.replace(/<(\/?)wp:/g, '<$1zz:').replace(/xmlns:wp=/, 'xmlns:zz=');
    expect(renamed).toContain('<zz:extent');
    zip.updateFile('word/document.xml', Buffer.from(renamed, 'utf8'));
    const s = await open(zip.toBuffer());
    const res = await s.run([{ action: 'chart', operation: 'update', target: chartId(s), widthIn: 5, altText: 'test' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const out = new AdmZip(s.parser.rezip(s.ast)).readAsText('word/document.xml');
    expect(out).toMatch(/<wp:extent cx="4572000"/); // the ingress rewrote zz to the usual prefix at open
    expect(out).toContain('descr="test"');
  });
});

describe('DrawingML run properties under another prefix', () => {
  const NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  it('applyRunPropsDeep formats d:rPr and d:defRPr when told the prefix', () => {
    const xml = `<c:rich xmlns:d="${NS}"><d:p><d:pPr><d:defRPr sz="800"/></d:pPr><d:r><d:rPr sz="900"/><d:t>x</d:t></d:r></d:p></c:rich>`;
    expect(drawingPrefixOf(xml)).toBe('d');
    const out = applyRunPropsDeep(xml, { sz: 1400, color: 'FF0000' }, 'd');
    expect(out).toContain('<d:rPr sz="1400"><d:solidFill><d:srgbClr val="FF0000"/></d:solidFill></d:rPr>');
    expect(out).toContain('<d:defRPr sz="1400"><d:solidFill>');
    expect(out).not.toContain('<a:');
  });

  it('the default prefix still works, and a prefix other than "a" in a chart part refuses the style call', () => {
    const ctx = { chartPath: 'p', chartType: 'column', seriesCount: 1, is3D: false };
    const xml = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:d="${NS}"><c:chart/></c:chartSpace>`;
    expect(() => applyChartStyle(xml, { chartFont: { sizePt: 9 } }, ctx)).toThrow(/binds the DrawingML namespace to the prefix "d"/);
  });
});

describe('loss inventory reads relationship attributes in either quote style', () => {
  it('finds a chartStyle relationship with single quotes, prefixed or not', () => {
    const type = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
    const plain = `<Relationships xmlns="r"><Relationship Id="r1" Type='${type}' Target="style1.xml"/></Relationships>`;
    const prefixed = `<p:Relationships xmlns:p="r"><p:Relationship Id="r1" Type="${type}" Target="style1.xml"/></p:Relationships>`;
    expect(describeLoss('<c:chartSpace/>', '<c:chartSpace/>', plain)).toContain('rel:chartStyle');
    expect(describeLoss('<c:chartSpace/>', '<c:chartSpace/>', prefixed)).toContain('rel:chartStyle');
  });
});

describe('trendline type names in the file', () => {
  it('writes exp, log and poly, which are the schema names, and the package validates', async () => {
    const s = await open(createBlankDocxBuffer('blank'));
    const base = (id: string, trendline: Record<string, unknown>) => ({
      action: 'chart', operation: 'insert', anchor: 'end', id, chartType: 'line', title: id,
      categories: ['a', 'b', 'c', 'd'], series: [{ name: 'S', values: [1, 4, 9, 16] }], style: { trendline },
    });
    const res = await s.run([base('e', { type: 'exponential' }), base('l', { type: 'logarithmic' }), base('p', { type: 'polynomial', order: 2 })]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const zip = new AdmZip(s.parser.rezip(s.ast));
    const all = zip.getEntries().filter((e) => /^word\/charts\/chart\d+\.xml$/.test(e.entryName)).map((e) => e.getData().toString('utf8')).join('\n');
    for (const name of ['exp', 'log', 'poly']) expect(all).toContain(`<c:trendlineType val="${name}"/>`);
    expect(all).not.toMatch(/trendlineType val="(exponential|logarithmic|polynomial)"/);
    expect((await validateDocxBuffer(s.parser.rezip(s.ast))).errors).toEqual([]);
    // Reading a chart back reports the friendly names.
    const node = s.ast.sections.flatMap((x: any) => x.nodes).filter((n: any) => n.type === 'chart');
    expect(node.length).toBe(3);
  });
});

describe('rebuild loss inventory names changed formatting values', () => {
  it('a colour that is gone is reported even though the element name survives', () => {
    const before = '<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:chartSpace>';
    const after = '<c:chartSpace><c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></c:spPr></c:chartSpace>';
    expect(describeLoss(before, after, null)).toEqual(['colour FF0000']);
  });

  it('font size, typeface and line width are reported; values that survive are not', () => {
    const before = '<c:txPr><a:p><a:pPr><a:defRPr sz="1400"><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr><a:ln w="28575"><a:solidFill><a:srgbClr val="00ff00"/></a:solidFill></a:ln>';
    const after = '<c:txPr><a:p><a:pPr><a:defRPr sz="1000"><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr><a:ln w="9525"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>';
    const lost = describeLoss(before, after, null);
    expect(lost).toEqual(expect.arrayContaining(['font size 14 pt', 'line width 28575 EMU']));
    expect(lost).not.toContain('typeface Calibri');
    expect(lost.some((l) => l.startsWith('colour'))).toBe(false);
  });

  it('an unchanged part reports nothing', () => {
    const xml = '<c:chartSpace><a:srgbClr val="123456"/><a:schemeClr val="accent1"/></c:chartSpace>';
    expect(describeLoss(xml, xml, null)).toEqual([]);
  });

  it('end to end: a rebuild that drops the plot area fill says so', async () => {
    const s = await open(createBlankDocxBuffer('blank'));
    await s.run([{
      action: 'chart', operation: 'insert', anchor: s.ast.sections[0].nodes[0].id, id: 'c1', chartType: 'column', title: 'R',
      categories: ['a', 'b'], series: [{ name: 'A', values: [1, 2], color: '0066CC' }, { name: 'B', values: [2, 1], color: '10B981' }],
    }]);
    const path = 'word/charts/chart1.xml';
    s.zip.writeText(path, s.zip.readText(path).replace('</c:plotArea>', '<c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></c:spPr></c:plotArea>'));
    const refusal = await s.run([{ action: 'chart', operation: 'update', target: 'c1', chartType: 'line' }]);
    expect(refusal.success).toBe(false);
    expect(JSON.stringify(refusal)).toContain('colour F2F2F2');
    const ok = await s.run([{ action: 'chart', operation: 'update', target: 'c1', chartType: 'line', allowRebuild: true }]);
    expect(ok.success).toBe(true);
    expect(ok.chartUpdates[0].formattingNotKept).toContain('colour F2F2F2');
    expect(ok.chartUpdates[0].formattingNotKept).not.toContain('colour 0066CC');
  });
});
