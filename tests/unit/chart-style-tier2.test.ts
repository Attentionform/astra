/** 0.5.3 chart style, Tier 2: chartFont, bars, lines, chartArea/plotArea, axis label settings. */
import { describe, expect, it } from 'vitest';
import { DocxParser, OperationEngine, createBlankDocxBuffer, parseStyleGuide, validateDocxBuffer } from '@attentionform/astra-engine';

async function open(guideLines?: string) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const guide = guideLines === undefined ? undefined : parseStyleGuide(`# Style Guide: t\n\n## Charts\n${guideLines}`).config;
  const engine = new OperationEngine(ast, parser, guide ? { styleGuide: guide as any } : undefined);
  const zip = parser.getZipAdapter();
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  return { parser, ast, zip, run };
}
type Session = Awaited<ReturnType<typeof open>>;
const xmlOf = (s: Session) => s.zip.readText('word/charts/chart1.xml');
const insertOp = (s: Session, extra: Record<string, unknown> = {}) => ({
  action: 'chart', operation: 'insert', anchor: s.ast.sections[0].nodes[0].id, id: 'c1', chartType: 'column', title: 'Revenue',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [{ name: 'A', values: [1, 2, 3, 4], color: '0066CC' }, { name: 'B', values: [4, 3, 2, 1], color: '10B981' }],
  ...extra,
});
const update = (extra: Record<string, unknown>) => ({ action: 'chart', operation: 'update', target: 'c1', ...extra });
const valid = async (s: Session) => (await validateDocxBuffer(s.parser.rezip(s.ast))).ok;
async function refused(s: Session, ops: unknown[], re: RegExp) {
  const before = xmlOf(s);
  const res = await s.run(ops);
  expect(res.success).toBe(false);
  expect(JSON.stringify(res)).toMatch(re);
  expect(xmlOf(s)).toBe(before);
}
async function withChart(extra: Record<string, unknown> = {}) {
  const s = await open();
  expect((await s.run([insertOp(s, extra)])).success).toBe(true);
  return s;
}

describe('chartFont', () => {
  it('sets family and colour everywhere and size everywhere except the title', async () => {
    const s = await withChart();
    const res = await s.run([update({ style: { chartFont: { family: 'Georgia', sizePt: 9, color: '334155' } } })]);
    expect(res.success).toBe(true);
    const xml = xmlOf(s);
    expect(xml).toContain('<a:latin typeface="Georgia"/>');
    expect(xml).toContain('334155');
    const title = /<c:chart>\s*<c:title>[\s\S]*?<\/c:title>/.exec(xml)![0];
    expect(title).toContain('sz="1400"');
    expect(title).toContain('Georgia');
    expect(/<c:txPr>[\s\S]*?sz="900"/.test(xml.slice(xml.lastIndexOf('</c:chart>')))).toBe(true);
    expect(await valid(s)).toBe(true);
  });

  it('a later titleFormat in the same call wins for the title', async () => {
    const s = await withChart();
    await s.run([update({ style: { chartFont: { sizePt: 9 }, titleFormat: { sizePt: 20 } } })]);
    expect(/<c:chart>\s*<c:title>[\s\S]*?<\/c:title>/.exec(xmlOf(s))![0]).toContain('sz="2000"');
  });

  it('refuses empty, bad size and bad colour', async () => {
    const s = await withChart();
    await refused(s, [update({ style: { chartFont: {} } })], /at least one of/);
    await refused(s, [update({ style: { chartFont: { sizePt: 0 } } })], /sizePt/);
    await refused(s, [update({ style: { chartFont: { color: 'zz' } } })], /not a colour/);
  });

  it('guide font family and size apply to a new chart; the call overrides them', async () => {
    const s = await open('- Font Family: Georgia\n- Font Size: 10 pt\n');
    await s.run([insertOp(s)]);
    expect(xmlOf(s)).toContain('Georgia');
    await s.run([update({ style: { chartFont: { family: 'Arial' } } })]);
    expect(xmlOf(s)).toContain('Arial');
    expect(xmlOf(s)).not.toContain('Georgia');
    expect(await valid(s)).toBe(true);
  });
});

describe('bars', () => {
  it('sets gap and overlap in schema order and replaces existing values', async () => {
    const s = await withChart();
    expect((await s.run([update({ style: { bars: { gapPercent: 40, overlapPercent: -10 } } })])).success).toBe(true);
    let xml = xmlOf(s);
    expect(xml).toContain('<c:gapWidth val="40"/>');
    expect(xml).toContain('<c:overlap val="-10"/>');
    expect(xml.indexOf('<c:gapWidth')).toBeLessThan(xml.indexOf('<c:overlap'));
    expect(xml.indexOf('<c:overlap')).toBeLessThan(xml.indexOf('<c:axId'));
    await s.run([update({ style: { bars: { gapPercent: 80 } } })]);
    xml = xmlOf(s);
    expect(xml.match(/<c:gapWidth/g)).toHaveLength(1);
    expect(xml).toContain('<c:gapWidth val="80"/>');
    expect(await valid(s)).toBe(true);
  });

  it('refuses range errors, charts with no bars, and overlap on 3D bars', async () => {
    const s = await withChart();
    await refused(s, [update({ style: { bars: { gapPercent: 501 } } })], /gapPercent/);
    await refused(s, [update({ style: { bars: { overlapPercent: 101 } } })], /overlapPercent/);
    const l = await withChart({ chartType: 'line' });
    await refused(l, [update({ style: { bars: { gapPercent: 10 } } })], /no|none/);
    const t = await withChart({ chartType: 'column3D' });
    await refused(t, [update({ style: { bars: { overlapPercent: 10 } } })], /3D/);
    expect((await t.run([update({ style: { bars: { gapPercent: 10 } } })])).success).toBe(true);
    expect(await valid(t)).toBe(true);
  });
});

describe('lines', () => {
  it('sets width, smoothing and markers on every line series', async () => {
    const s = await withChart({ chartType: 'line' });
    expect((await s.run([update({ style: { lines: { widthPt: 3, smooth: true, showMarkers: false } } })])).success).toBe(true);
    const sers = xmlOf(s).match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    for (const ser of sers) {
      expect(ser).toContain('w="38100"');
      expect(ser).toContain('<c:smooth val="1"/>');
      expect(ser).toContain('<c:symbol val="none"/>');
    }
    expect(await valid(s)).toBe(true);
    await s.run([update({ style: { lines: { showMarkers: true, smooth: false } } })]);
    const after = xmlOf(s).match(/<c:ser>[\s\S]*?<\/c:ser>/g)!;
    expect(after[0]).toContain('<c:symbol val="circle"/>');
    expect(after[0]).toContain('<c:smooth val="0"/>');
    expect(await valid(s)).toBe(true);
  });

  it('keeps the group-level smooth of a combo chart equal to the series', async () => {
    const s = await withChart({
      chartType: 'combo',
      series: [{ name: 'R', values: [1, 2, 3, 4], chartType: 'column' }, { name: 'M', values: [4, 3, 2, 1], chartType: 'line' }],
    });
    await s.run([update({ style: { lines: { smooth: true } } })]);
    const line = /<c:lineChart>[\s\S]*?<\/c:lineChart>/.exec(xmlOf(s))![0];
    expect(line).toContain('<c:smooth val="1"/></c:ser>');
    expect(line.match(/<c:smooth val="1"\/>/g)).toHaveLength(2);
    expect(await valid(s)).toBe(true);
  });

  it('refuses a chart with no line group and bad values', async () => {
    const s = await withChart();
    await refused(s, [update({ style: { lines: { widthPt: 2 } } })], /no line|none|has none/);
    const l = await withChart({ chartType: 'line' });
    await refused(l, [update({ style: { lines: { widthPt: 0 } } })], /widthPt/);
  });
});

describe('chartArea and plotArea', () => {
  it('fills and borders both areas, in valid order', async () => {
    const s = await withChart();
    const res = await s.run([update({ style: { chartArea: { fillColor: 'F8FAFC', borderColor: 'none' }, plotArea: { fillColor: 'FFFFFF', borderColor: 'CBD5E1' } } })]);
    expect(res.success).toBe(true);
    const xml = xmlOf(s);
    const tail = xml.slice(xml.lastIndexOf('</c:chart>'));
    expect(tail).toContain('F8FAFC');
    expect(tail).toContain('<a:noFill/>');
    const plot = /<c:plotArea>[\s\S]*<\/c:plotArea>/.exec(xml)![0];
    expect(plot).toContain('CBD5E1');
    expect(plot.lastIndexOf('<c:spPr>')).toBeGreaterThan(plot.lastIndexOf('</c:valAx>'));
    expect(await valid(s)).toBe(true);
    await s.run([update({ style: { chartArea: { fillColor: 'none' } } })]);
    expect(await valid(s)).toBe(true);
  });

  it('refuses bad colours and empty objects', async () => {
    const s = await withChart();
    await refused(s, [update({ style: { chartArea: { fillColor: 'bad' } } })], /not a colour/);
    await refused(s, [update({ style: { plotArea: {} } })], /at least one of/);
  });
});

describe('axis label settings', () => {
  it('labelSizePt and labelRotationDeg write txPr in axis order', async () => {
    const s = await withChart();
    const res = await s.run([update({ style: { axes: { category: { labelSizePt: 8, labelRotationDeg: -45 }, value: { labelSizePt: 9 } } } })]);
    expect(res.success).toBe(true);
    const cat = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xmlOf(s))![0];
    expect(cat).toContain('sz="800"');
    expect(cat).toContain('rot="-2700000"');
    expect(cat.indexOf('<c:txPr>')).toBeLessThan(cat.indexOf('<c:crossAx'));
    expect(/<c:valAx>[\s\S]*?<\/c:valAx>/.exec(xmlOf(s))![0]).toContain('sz="900"');
    expect(await valid(s)).toBe(true);
    await s.run([update({ style: { axes: { category: { labelRotationDeg: 30, labelSizePt: 11 } } } })]);
    const cat2 = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xmlOf(s))![0];
    expect(cat2).toContain('rot="1800000"');
    expect(cat2).toContain('sz="1100"');
    expect(await valid(s)).toBe(true);
  });

  it('showAxisLine draws or hides the axis line', async () => {
    const s = await withChart();
    await s.run([update({ style: { axes: { category: { showAxisLine: false }, value: { showAxisLine: true } } } })]);
    const cat = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xmlOf(s))![0];
    expect(/<c:spPr>[\s\S]*?<a:ln[^>]*>\s*<a:noFill\/>/.test(cat)).toBe(true);
    const val = /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(xmlOf(s))![0];
    expect(val).toContain('BFBFBF');
    expect(await valid(s)).toBe(true);
    await s.run([update({ style: { axes: { category: { showAxisLine: true } } } })]);
    expect(/<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xmlOf(s))![0]).toContain('BFBFBF');
    expect(await valid(s)).toBe(true);
  });

  it('refuses out-of-range values and pie charts', async () => {
    const s = await withChart();
    await refused(s, [update({ style: { axes: { category: { labelRotationDeg: 91 } } } })], /labelRotationDeg/);
    await refused(s, [update({ style: { axes: { value: { labelSizePt: 0 } } } })], /labelSizePt/);
    const p = await withChart({ chartType: 'pie', series: [{ name: 'A', values: [1, 2, 3, 4] }] });
    await refused(p, [update({ style: { axes: { value: { showAxisLine: true } } } })], /no value axis/);
  });
});
