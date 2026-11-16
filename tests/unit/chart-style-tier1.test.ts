/**
 * 0.5.3 chart style, Tier 1: titleFormat, removeTitle, gridlineColor, per-point colours,
 * trendline.seriesNumbers.
 */
import { describe, expect, it } from 'vitest';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer, validateOperationStrict } from '@attentionform/astra-engine';

async function open() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const zip = parser.getZipAdapter();
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  return { parser, ast, zip, run };
}
type Session = Awaited<ReturnType<typeof open>>;
const CHART = 'word/charts/chart1.xml';
const xmlOf = (s: Session) => s.zip.readText(CHART);
const insertOp = (s: Session, extra: Record<string, unknown> = {}) => ({
  action: 'chart', operation: 'insert', anchor: s.ast.sections[0].nodes[0].id, id: 'c1', chartType: 'column', title: 'Revenue',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [{ name: 'A', values: [1, 2, 3, 4], color: '0066CC' }, { name: 'B', values: [4, 3, 2, 1], color: '10B981' }],
  ...extra,
});
const update = (extra: Record<string, unknown>) => ({ action: 'chart', operation: 'update', target: 'c1', ...extra });
const valid = async (s: Session) => (await validateDocxBuffer(s.parser.rezip(s.ast))).ok;
const titleBlock = (xml: string) => /<c:chart>\s*<c:title>[\s\S]*?<\/c:title>/.exec(xml)![0];
async function refused(s: Session, ops: unknown[], re: RegExp) {
  const before = xmlOf(s);
  const res = await s.run(ops);
  expect(res.success).toBe(false);
  expect(JSON.stringify(res)).toMatch(re);
  expect(xmlOf(s)).toBe(before);
}

describe('style.titleFormat', () => {
  it('changes size, bold and colour of the existing title and keeps the text', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ style: { titleFormat: { sizePt: 12, bold: false, color: '1E293B' } } })]);
    expect(res.success).toBe(true);
    const title = titleBlock(xmlOf(s));
    expect(title).toContain('sz="1200"');
    expect(title).toContain('b="0"');
    expect(title).toContain('<a:srgbClr val="1E293B"/>');
    expect(title).toContain('>Revenue<');
    expect(await valid(s)).toBe(true);
  });

  it('only the given properties change', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await s.run([update({ style: { titleFormat: { color: 'FF0000' } } })]);
    const title = titleBlock(xmlOf(s));
    expect(title).toContain('sz="1400"');
    expect(title).toContain('b="1"');
    expect(title).toContain('FF0000');
  });

  it('works on insert together with title, and is refused when the chart has no title', async () => {
    const s = await open();
    const res = await s.run([insertOp(s, { style: { titleFormat: { sizePt: 20 } } })]);
    expect(res.success).toBe(true);
    expect(titleBlock(xmlOf(s))).toContain('sz="2000"');
    const t = await open();
    await t.run([insertOp(t)]);
    await t.run([update({ removeTitle: true })]);
    await refused(t, [update({ style: { titleFormat: { sizePt: 12 } } })], /no title/);
  });

  it('refuses empty, out-of-range and bad values', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await refused(s, [update({ style: { titleFormat: {} } })], /at least one of/);
    await refused(s, [update({ style: { titleFormat: { sizePt: 0 } } })], /sizePt/);
    await refused(s, [update({ style: { titleFormat: { sizePt: 500 } } })], /sizePt/);
    await refused(s, [update({ style: { titleFormat: { color: 'nope' } } })], /colour|color/);
  });

  it('a title text change keeps the format', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await s.run([update({ style: { titleFormat: { sizePt: 18 } } })]);
    await s.run([update({ title: 'Sales' })]);
    const title = titleBlock(xmlOf(s));
    expect(title).toContain('>Sales<');
    expect(title).toContain('sz="1800"');
  });
});

describe('removeTitle', () => {
  it('removes the title element, shows no automatic title, and clears the metadata title', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ removeTitle: true })]);
    expect(res.success).toBe(true);
    const xml = xmlOf(s);
    expect(xml).not.toMatch(/<c:chart>\s*<c:title>/);
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>');
    expect(await valid(s)).toBe(true);
    const node = s.ast.sections[0].nodes.find((n: any) => n.type === 'chart');
    expect(node.metadata.title).toBeUndefined();
  });

  it('is refused together with title or titleFormat, and on insert', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await refused(s, [update({ removeTitle: true, title: 'x' })], /cannot be combined/);
    await refused(s, [update({ removeTitle: true, style: { titleFormat: { sizePt: 10 } } })], /cannot be combined/);
    const t = await open();
    const res = await t.run([insertOp(t, { removeTitle: true })]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/for operation:'update'/);
  });

  it('a title can be set again afterwards', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await s.run([update({ removeTitle: true })]);
    await s.run([update({ title: 'Back' })]);
    expect(xmlOf(s)).toMatch(/<c:chart>\s*<c:title>/);
    expect(xmlOf(s)).toContain('>Back<');
    expect(xmlOf(s)).toContain('<c:autoTitleDeleted val="0"/>');
  });
});

describe('axes gridlineColor', () => {
  it('recolours existing gridlines and keeps their width', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ style: { axes: { value: { gridlineColor: 'FF8800' } } } })]);
    expect(res.success).toBe(true);
    const grid = /<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/.exec(xmlOf(s))![0];
    expect(grid).toContain('FF8800');
    expect(grid).toContain('w="9525"');
    expect(grid).not.toContain('D9D9D9');
  });

  it('is refused while gridlines are off, and works with majorGridlines:true in the same call', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await refused(s, [update({ style: { axes: { category: { gridlineColor: 'FF8800' } } } })], /gridlines.*off|off on this axis/);
    const res = await s.run([update({ style: { axes: { category: { majorGridlines: true, gridlineColor: 'FF8800' } } } })]);
    expect(res.success).toBe(true);
    const cat = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xmlOf(s))![0];
    expect(cat).toContain('FF8800');
    await refused(s, [update({ style: { axes: { value: { majorGridlines: false, gridlineColor: 'FF8800' } } } })], /conflicts/);
  });

  it('is refused on a pie chart (no axes) and for a bad colour', async () => {
    const s = await open();
    await s.run([insertOp(s, { chartType: 'pie', series: [{ name: 'A', values: [1, 2, 3, 4] }] })]);
    await refused(s, [update({ style: { axes: { value: { gridlineColor: 'FF8800' } } } })], /no value axis/);
    const t = await open();
    await t.run([insertOp(t)]);
    await refused(t, [update({ style: { axes: { value: { gridlineColor: 'zzz' } } } })], /not a colour/);
  });
});

describe('style.series points', () => {
  it('colours single bars, creating dPt in index order and replacing an existing one', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 3, color: 'F59E0B' }, { pointNumber: 1, color: '4F46E5' }] }] } })]);
    expect(res.success).toBe(true);
    let ser = /<c:ser>[\s\S]*?<\/c:ser>/.exec(xmlOf(s))![0];
    const dPts = ser.match(/<c:dPt>[\s\S]*?<\/c:dPt>/g)!;
    expect(dPts).toHaveLength(2);
    expect(dPts[0]).toContain('<c:idx val="0"/>');
    expect(dPts[0]).toContain('4F46E5');
    expect(dPts[1]).toContain('<c:idx val="2"/>');
    expect(dPts[1]).toContain('F59E0B');
    expect(ser.indexOf('<c:dPt>')).toBeLessThan(ser.indexOf('<c:cat>'));
    await s.run([update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 3, color: '00FF00' }] }] } })]);
    ser = /<c:ser>[\s\S]*?<\/c:ser>/.exec(xmlOf(s))![0];
    expect(ser.match(/<c:dPt>/g)).toHaveLength(2);
    expect(ser).toContain('00FF00');
    expect(ser).not.toContain('F59E0B');
    expect(await valid(s)).toBe(true);
  });

  it('colours pie slices and line markers', async () => {
    const s = await open();
    await s.run([insertOp(s, { chartType: 'pie', series: [{ name: 'A', values: [1, 2, 3, 4] }] })]);
    const res = await s.run([update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 2, color: 'AA00AA' }] }] } })]);
    expect(res.success).toBe(true);
    const dPt = (xmlOf(s).match(/<c:dPt>[\s\S]*?<\/c:dPt>/g) ?? []).filter((d) => d.includes('<c:idx val="1"/>'));
    expect(dPt.join('')).toContain('AA00AA');
    expect(dPt.join('')).not.toContain('FF0000');
    const t = await open();
    await t.run([insertOp(t, { chartType: 'line' })]);
    await t.run([update({ style: { series: [{ seriesNumber: 2, points: [{ pointNumber: 4, color: 'AA00AA' }] }] } })]);
    const ser2 = xmlOf(t).match(/<c:ser>[\s\S]*?<\/c:ser>/g)![1];
    expect(ser2).toContain('<c:marker>');
    expect(ser2).toContain('AA00AA');
    expect(await valid(t)).toBe(true);
  });

  it('refuses a number outside the range with the range, duplicates, 3D, area and bad colours', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    await refused(s, [update({ style: { series: [{ seriesNumber: 3, points: [{ pointNumber: 1, color: 'FF0000' }] }] } })], /outside 1\.\.2/);
    await refused(s, [update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 5, color: 'FF0000' }] }] } })], /outside 1\.\.4/);
    await refused(s, [update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 1, color: 'FF0000' }, { pointNumber: 1, color: '00FF00' }] }] } })], /twice/);
    await refused(s, [update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 1, color: 'FF0000' }] }, { seriesNumber: 1, points: [{ pointNumber: 2, color: 'FF0000' }] }] } })], /twice/);
    await refused(s, [update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 1, color: 'bad' }] }] } })], /not a colour/);
    const t = await open();
    await t.run([insertOp(t, { chartType: 'column3D' })]);
    await refused(t, [update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 1, color: 'FF0000' }] }] } })], /3D/);
    const a = await open();
    await a.run([insertOp(a, { chartType: 'area' })]);
    await refused(a, [update({ style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 1, color: 'FF0000' }] }] } })], /area/);
  });

  it('series numbers mean the chart before the call when removeSeries runs in the same call', async () => {
    const s = await open();
    await s.run([insertOp(s, { series: [
      { name: 'A', values: [1, 2, 3, 4] }, { name: 'B', values: [4, 3, 2, 1] }, { name: 'C', values: [2, 2, 2, 2] },
    ] })]);
    const res = await s.run([update({ removeSeries: ['A'], allowRebuild: true, style: { series: [{ seriesNumber: 3, points: [{ pointNumber: 1, color: 'AA00AA' }] }] } })]);
    expect(res.success).toBe(true);
    const sers = xmlOf(s).match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(sers).toHaveLength(2);
    expect(sers[1]).toContain('AA00AA');
    expect(sers[0]).not.toContain('AA00AA');
    await refused(s, [update({ removeSeries: ['B'], allowRebuild: true, style: { series: [{ seriesNumber: 1, points: [{ pointNumber: 1, color: 'AA00AA' }] }] } })], /removes in the same call|names series/);
  });
});

describe('trendline.seriesNumbers', () => {
  it('applies to the numbered series only, and the old key is refused', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ style: { trendline: { type: 'linear', seriesNumbers: [2] } } })]);
    expect(res.success).toBe(true);
    const sers = xmlOf(s).match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(sers[0]).not.toContain('<c:trendline>');
    expect(sers[1]).toContain('<c:trendline>');
    expect(() => validateOperationStrict(update({ style: { trendline: { type: 'linear', series: [1] } } }))).toThrow();
    expect(() => validateOperationStrict(update({ style: { trendline: { type: 'linear', seriesNumbers: [1] } } }))).not.toThrow();
  });
});

describe('guide chart defaults', () => {
  async function guided(lines: string) {
    const { parseStyleGuide } = await import('@attentionform/astra-engine');
    const guide = parseStyleGuide(`# Style Guide: t\n\n## Charts\n${lines}`).config;
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const engine = new OperationEngine(ast, parser, { styleGuide: guide as any });
    const zip = parser.getZipAdapter();
    const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
    return { parser, ast, zip, run } as Session;
  }

  it('gridline colour and title size from the guide apply to a new chart', async () => {
    const s = await guided('- Gridline Color: 112233\n- Title Size: 18 pt\n');
    const res = await s.run([insertOp(s)]);
    expect(res.success).toBe(true);
    const xml = xmlOf(s);
    expect(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/.exec(xml)![0]).toContain('112233');
    expect(xml).not.toContain('D9D9D9');
    expect(titleBlock(xml)).toContain('sz="1800"');
    expect(await valid(s)).toBe(true);
  });

  it('an explicit style key beats the guide', async () => {
    const s = await guided('- Gridline Color: 112233\n- Title Size: 18 pt\n');
    await s.run([insertOp(s, { style: { titleFormat: { sizePt: 10 }, axes: { value: { gridlineColor: 'ABCDEF' } } } })]);
    const xml = xmlOf(s);
    expect(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/.exec(xml)![0]).toContain('ABCDEF');
    expect(titleBlock(xml)).toContain('sz="1000"');
  });

  it('a pie chart (no gridlines) ignores the gridline default silently', async () => {
    const s = await guided('- Gridline Color: 112233\n');
    const res = await s.run([insertOp(s, { chartType: 'pie', series: [{ name: 'A', values: [1, 2, 3, 4] }] })]);
    expect(res.success).toBe(true);
    expect(xmlOf(s)).not.toContain('112233');
  });

  it('without a guide the chart keeps D9D9D9 gridlines', async () => {
    const s = await open();
    await s.run([insertOp(s)]);
    expect(xmlOf(s)).toContain('D9D9D9');
  });
});
