/**
 * 0.5.3 (DN1): structural chart updates keep the chart part.
 *   Case A  add series at the end        Case B  removeSeries
 *   Case C  type change inside the column/bar family
 *   Case D  anything else needs allowRebuild:true and reports what it did not keep
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const WORD_CHART = join(__dirname, '..', 'fixtures', 'word-chart-with-workbook.docx');

async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const zip = parser.getZipAdapter();
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  return { parser, ast, zip, run };
}
type Session = Awaited<ReturnType<typeof open>>;

const blank = () => open(createBlankDocxBuffer('blank'));
const CHART = 'word/charts/chart1.xml';
const chartXml = (s: Session) => s.zip.readText(CHART);
const sheetOf = (s: Session): string => {
  const path = s.zip.entryNames!().find((n) => /^word\/embeddings\/.*\.xlsx$/.test(n))!;
  return new AdmZip(s.zip.readBuffer!(path) as Buffer).readAsText('xl/worksheets/sheet1.xml');
};
const insertOp = (s: Session, extra: Record<string, unknown> = {}) => ({
  action: 'chart', operation: 'insert', anchor: s.ast.sections[0].nodes[0].id, id: 'c1', chartType: 'column', title: 'Revenue',
  categories: ['Q1', 'Q2', 'Q3'],
  series: [
    { name: 'A', values: [1, 2, 3], color: '0066CC' },
    { name: 'B', values: [3, 2, 1], color: '10B981' },
    { name: 'C', values: [2, 2, 2], color: 'F59E0B' },
  ],
  ...extra,
});
const update = (extra: Record<string, unknown>) => ({ action: 'chart', operation: 'update', target: 'c1', ...extra });
const seriesBlocks = (xml: string): string[] => xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g) ?? [];
const nameOf = (ser: string): string => /<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/.exec(ser)![1];

/** Injects nodes the engine does not model, in schema order. */
function withUnmodelled(xml: string): string {
  let out = xml;
  out = out.replace(/(<c:ser>\s*<c:idx val="0"\/>[\s\S]*?)(<c:cat>)/, '$1<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></c:spPr></c:dPt>$2');
  out = out.replace(/(<c:ser>\s*<c:idx val="1"\/>[\s\S]*?)(<c:cat>)/, '$1<c:trendline><c:trendlineType val="linear"/></c:trendline>$2');
  out = out.replace(/(<c:ser>\s*<c:idx val="2"\/>[\s\S]*?)(<c:cat>)/, '$1<c:dLbls><c:dLbl><c:idx val="0"/><c:showVal val="1"/></c:dLbl><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>$2');
  out = out.replace('</c:plotArea>', '<c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill></c:spPr></c:plotArea>');
  return out;
}
async function styled(extra: Record<string, unknown> = {}) {
  const s = await blank();
  await s.run([insertOp(s, extra)]);
  s.zip.writeText(CHART, withUnmodelled(chartXml(s)));
  return s;
}
const validOf = async (s: Session) => (await validateDocxBuffer(s.parser.rezip(s.ast))).ok;

describe('case C: type change inside the column/bar family', () => {
  it('column to bar flips barDir and the two axis positions and keeps everything else', async () => {
    const s = await styled();
    const before = chartXml(s);
    const res = await s.run([update({ chartType: 'bar' })]);
    expect(res.success).toBe(true);
    const after = chartXml(s);
    expect(after).toContain('<c:barDir val="bar"/>');
    expect(after).toMatch(/<c:catAx>[\s\S]*<c:axPos val="l"\/>[\s\S]*<\/c:catAx>/);
    expect(after).toMatch(/<c:valAx>[\s\S]*<c:axPos val="b"\/>[\s\S]*<\/c:valAx>/);
    expect(after).toContain('<c:dPt>');
    expect(after).toContain('<c:trendline>');
    expect(after).toContain('F2F2F2');
    expect(after).toContain('<c:gapWidth val="150"/>');
    expect(res.chartUpdates[0]).toMatchObject({ updateMode: 'in-place', workbook: 'untouched' });
    expect(after.replace(/<c:barDir val="bar"\/>/, '<c:barDir val="col"/>').replace(/<c:axPos val="[lb]"\/>/g, '')).toBe(
      before.replace(/<c:axPos val="[lb]"\/>/g, '')
    );
    expect(await validOf(s)).toBe(true);
  });

  it('clustered to stacked sets grouping and overlap 100; back removes the overlap it added', async () => {
    const s = await styled();
    await s.run([update({ chartType: 'columnStacked' })]);
    let xml = chartXml(s);
    expect(xml).toContain('<c:grouping val="stacked"/>');
    expect(xml).toContain('<c:overlap val="100"/>');
    expect(xml).toContain('<c:gapWidth val="150"/>');
    // 'column' on a stacked column chart keeps the stacked form (the base-family rule); un-stack through the bar family.
    await s.run([update({ chartType: 'bar' })]);
    await s.run([update({ chartType: 'column' })]);
    xml = chartXml(s);
    expect(xml).toContain('<c:grouping val="clustered"/>');
    expect(xml).toContain('<c:barDir val="col"/>');
    expect(xml).not.toContain('<c:overlap');
    expect(await validOf(s)).toBe(true);
  });

  it('an authored overlap other than 100 stays when going back to clustered', async () => {
    const s = await styled({ chartType: 'columnStacked' });
    s.zip.writeText(CHART, chartXml(s).replace('<c:overlap val="100"/>', '<c:overlap val="40"/>'));
    await s.run([update({ chartType: 'bar' })]);
    expect(chartXml(s)).toContain('<c:overlap val="40"/>');
  });

  it('percent stacked sets the percent number format only while the format is the default', async () => {
    const s = await styled();
    const res = await s.run([update({ chartType: 'columnPercentStacked' })]);
    expect(res.success).toBe(true);
    expect(chartXml(s)).toContain('formatCode="0%" sourceLinked="0"');
    expect(res.chartUpdates[0].styleNotApplied).toBeUndefined();
    await s.run([update({ chartType: 'bar' })]);
    expect(chartXml(s)).toContain('<c:numFmt formatCode="General" sourceLinked="1"/>');
    expect(chartXml(s)).toContain('<c:grouping val="clustered"/>');

    const custom = await styled();
    custom.zip.writeText(CHART, chartXml(custom).replace('<c:numFmt formatCode="General" sourceLinked="1"/>', '<c:numFmt formatCode="$#,##0" sourceLinked="0"/>'));
    const r2 = await custom.run([update({ chartType: 'columnPercentStacked' })]);
    expect(r2.success).toBe(true);
    expect(chartXml(custom)).toContain('formatCode="$#,##0"');
    expect(r2.chartUpdates[0].styleNotApplied?.join(' ')).toMatch(/number format/);
  });

  it('works on a chart made in Word (theme colours, extensions, style parts stay)', async () => {
    const s = await open(readFileSync(WORD_CHART));
    const id = s.ast.sections.flatMap((x: any) => x.nodes).find((n: any) => n.type === 'chart').id;
    const before = chartXml(s);
    const res = await s.run([{ action: 'chart', operation: 'update', target: id, chartType: 'barStacked' }]);
    expect(res.success).toBe(true);
    const after = chartXml(s);
    expect(after).toContain('<c:barDir val="bar"/>');
    expect(after).toContain('<c:grouping val="stacked"/>');
    expect(after).toContain('<c:overlap val="100"/>');
    expect(after.length).toBeGreaterThan(before.length - 50);
    expect(s.zip.hasFile('word/charts/style1.xml')).toBe(true);
    expect(s.zip.hasFile('word/charts/colors1.xml')).toBe(true);
  });

  it('a combo chart is refused without allowRebuild', async () => {
    const s = await blank();
    await s.run([insertOp(s, {
      chartType: 'combo',
      series: [
        { name: 'A', values: [1, 2, 3], chartType: 'column' },
        { name: 'B', values: [3, 2, 1], chartType: 'line', axis: 'secondary' },
      ],
    })]);
    const res = await s.run([update({ chartType: 'bar' })]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/allowRebuild/);
  });
});

describe('case A: add series in place', () => {
  it('appends series with the next idx and order, own colour, shifted columns, and keeps the rest', async () => {
    const s = await styled();
    const before = seriesBlocks(chartXml(s));
    const res = await s.run([update({ series: [
      { name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] },
      { name: 'D', values: [9, 8, 7], color: 'AA00AA' },
    ] })]);
    expect(res.success).toBe(true);
    const after = seriesBlocks(chartXml(s));
    expect(after).toHaveLength(4);
    expect(after.slice(0, 3)).toEqual(before);
    const added = after[3];
    expect(added).toContain('<c:idx val="3"/>');
    expect(added).toContain('<c:order val="3"/>');
    expect(added).toContain('AA00AA');
    expect(added).toContain('Sheet1!$E$1');
    expect(added).toContain('Sheet1!$E$2:$E$4');
    expect(added).not.toContain('<c:dPt>');
    expect(added).not.toContain('<c:trendline>');
    expect(added).toContain('<c:dLbls>');
    expect(added).not.toContain('<c:dLbl>');
    expect(nameOf(added)).toBe('D');
    expect(sheetOf(s)).toContain('>D<');
    expect(res.chartUpdates[0].updateMode).toBe('in-place');
    expect(chartXml(s)).toContain('F2F2F2');
    expect(await validOf(s)).toBe(true);
  });

  it('a new series without a colour takes the next palette colour', async () => {
    const s = await blank();
    await s.run([insertOp(s)]);
    await s.run([update({ series: [
      { name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] }, { name: 'D', values: [1, 1, 1] },
    ] })]);
    expect(seriesBlocks(chartXml(s))[3]).toContain('<a:srgbClr val="EF4444"/>');
  });

  it('adds to a line chart, keeping the marker shape with the new colour', async () => {
    const s = await blank();
    await s.run([insertOp(s, { chartType: 'line' })]);
    await s.run([update({ series: [
      { name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] }, { name: 'D', values: [1, 1, 1], color: '112233' },
    ] })]);
    const added = seriesBlocks(chartXml(s))[3];
    expect(added).toContain('112233');
    expect(await validOf(s)).toBe(true);
  });

  it('adds to a 3D column chart', async () => {
    const s = await blank();
    await s.run([insertOp(s, { chartType: 'column3D' })]);
    const res = await s.run([update({ series: [
      { name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] }, { name: 'D', values: [1, 1, 1] },
    ] })]);
    expect(res.success).toBe(true);
    expect(seriesBlocks(chartXml(s))).toHaveLength(4);
    expect(await validOf(s)).toBe(true);
  });

  it('adding a series to a pie chart needs allowRebuild', async () => {
    const s = await blank();
    await s.run([insertOp(s, { chartType: 'pie', series: [{ name: 'A', values: [1, 2, 3] }] })]);
    const res = await s.run([update({ series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [1, 1, 1] }] })]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/allowRebuild/);
  });

  it('data change on a Word workbook is refused before anything is written', async () => {
    const s = await open(readFileSync(WORD_CHART));
    const id = s.ast.sections.flatMap((x: any) => x.nodes).find((n: any) => n.type === 'chart').id;
    const before = chartXml(s);
    const res = await s.run([{ action: 'chart', operation: 'update', target: id, series: [
      { name: 'Series 1', values: [1, 2, 3, 4] }, { name: 'Series 2', values: [1, 2, 3, 4] }, { name: 'Series 3', values: [1, 2, 3, 4] }, { name: 'Series 4', values: [1, 2, 3, 4] },
    ] }]);
    expect(res.success).toBe(false);
    expect(chartXml(s)).toBe(before);
  });
});

describe('case B: removeSeries', () => {
  it('removes by name, renumbers, shifts columns and keeps the remaining series untouched', async () => {
    const s = await styled();
    const before = seriesBlocks(chartXml(s));
    const res = await s.run([update({ removeSeries: ['B'] })]);
    expect(res.success).toBe(true);
    const after = seriesBlocks(chartXml(s));
    expect(after.map(nameOf)).toEqual(['A', 'C']);
    expect(after[0]).toBe(before[0]);
    expect(after[1]).toContain('<c:idx val="1"/>');
    expect(after[1]).toContain('<c:order val="1"/>');
    expect(after[1]).toContain('Sheet1!$C$2:$C$4');
    expect(after[1]).toContain('<c:dLbls>');
    expect(chartXml(s)).not.toContain('<c:trendline>');
    expect(sheetOf(s)).not.toContain('>B<');
    expect(sheetOf(s)).toContain('>C<');
    expect(await validOf(s)).toBe(true);
  });

  it('removes first, last, by one-based position, and several at once', async () => {
    const a = await styled();
    await a.run([update({ removeSeries: [1] })]);
    expect(seriesBlocks(chartXml(a)).map(nameOf)).toEqual(['B', 'C']);
    const b = await styled();
    await b.run([update({ removeSeries: ['C'] })]);
    expect(seriesBlocks(chartXml(b)).map(nameOf)).toEqual(['A', 'B']);
    const c = await styled();
    await c.run([update({ removeSeries: ['A', 3] })]);
    expect(seriesBlocks(chartXml(c)).map(nameOf)).toEqual(['B']);
  });

  it('refuses a missing name, a bad position, a repeated entry, a duplicate name and removing everything', async () => {
    const s = await blank();
    await s.run([insertOp(s, { series: [
      { name: 'A', values: [1, 2, 3] }, { name: 'A', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] },
    ] })]);
    const before = chartXml(s);
    for (const [removeSeries, re] of [
      [['Nope'], /matches no series/],
      [[9], /outside 1\.\.3/],
      [[1, 1], /twice/],
      [['A'], /matches 2 series/],
      [[1, 2, 3], /every series/],
    ] as Array<[unknown[], RegExp]>) {
      const res = await s.run([update({ removeSeries })]);
      expect(res.success).toBe(false);
      expect(JSON.stringify(res)).toMatch(re);
    }
    expect(chartXml(s)).toBe(before);
  });

  it('refuses a shorter series array and names what it would drop', async () => {
    const s = await blank();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ series: [{ name: 'A', values: [1, 2, 3] }] })]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/would drop \\"B\\", \\"C\\"/);
    expect(JSON.stringify(res)).toMatch(/removeSeries/);
  });

  it('removeSeries together with series applies the series positionally to what remains', async () => {
    const s = await blank();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ removeSeries: ['A'], series: [{ name: 'B', values: [7, 7, 7] }, { name: 'C', values: [2, 2, 2] }] })]);
    expect(res.success).toBe(true);
    expect(seriesBlocks(chartXml(s)).map(nameOf)).toEqual(['B', 'C']);
    expect(chartXml(s)).toContain('<c:v>7</c:v>');
  });

  it('removes a series and fixes legend entries', async () => {
    const s = await blank();
    await s.run([insertOp(s)]);
    s.zip.writeText(CHART, chartXml(s).replace(/(<c:legend>\s*<c:legendPos val="[a-z]+"\/>)/, '$1<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry>'));
    expect(chartXml(s)).toContain('<c:legendEntry>');
    await s.run([update({ removeSeries: ['B'] })]);
    const xml = chartXml(s);
    expect(xml.match(/<c:legendEntry>/g)).toHaveLength(1);
    expect(xml).toMatch(/<c:legendEntry><c:idx val="1"\/>/);
  });

  it('removes from a 3D bar chart and from a line chart', async () => {
    for (const chartType of ['bar3D', 'line']) {
      const s = await blank();
      await s.run([insertOp(s, { chartType })]);
      const res = await s.run([update({ removeSeries: ['B'] })]);
      expect(res.success).toBe(true);
      expect(seriesBlocks(chartXml(s)).map(nameOf)).toEqual(['A', 'C']);
      expect(await validOf(s)).toBe(true);
    }
  });

  it('removal on a Word-authored workbook is a data change: refused without allowRebuild', async () => {
    const s = await open(readFileSync(WORD_CHART));
    const id = s.ast.sections.flatMap((x: any) => x.nodes).find((n: any) => n.type === 'chart').id;
    const before = chartXml(s);
    const res = await s.run([{ action: 'chart', operation: 'update', target: id, removeSeries: [3] }]);
    expect(res.success).toBe(false);
    expect(JSON.stringify(res)).toMatch(/allowRebuild/);
    expect(chartXml(s)).toBe(before);
    const ok = await s.run([{ action: 'chart', operation: 'update', target: id, removeSeries: [3], allowRebuild: true }]);
    expect(ok.success).toBe(true);
    expect(seriesBlocks(chartXml(s))).toHaveLength(2);
  });
});

describe('case D: a rebuild needs allowRebuild and says what it did not keep', () => {
  it('cross-family change is refused with an inventory, nothing changes', async () => {
    const s = await styled();
    const before = chartXml(s);
    const res = await s.run([update({ chartType: 'line' })]);
    expect(res.success).toBe(false);
    const text = JSON.stringify(res);
    expect(text).toMatch(/allowRebuild/);
    expect(text).toMatch(/dPt/);
    expect(text).toMatch(/trendline/);
    expect(text).toMatch(/cannot read every extension/);
    expect(chartXml(s)).toBe(before);
  });

  it('with allowRebuild the part is rebuilt, the modelled style is re-applied and the loss is reported', async () => {
    const s = await styled();
    s.zip.writeText(CHART, chartXml(s).replace(/<c:legendPos val="[a-z]+"\/>/, '<c:legendPos val="t"/>'));
    const res = await s.run([update({ chartType: 'line', allowRebuild: true })]);
    expect(res.success).toBe(true);
    const report = res.chartUpdates[0];
    expect(report.updateMode).toBe('rebuilt');
    expect(report.formattingNotKept).toEqual(expect.arrayContaining(['dPt', 'trendline']));
    const xml = chartXml(s);
    expect(xml).toContain('<c:lineChart>');
    expect(xml).toContain('<c:legendPos val="t"/>');
    expect(xml).not.toContain('<c:dPt>');
    expect(await validOf(s)).toBe(true);
  });

  it('style that does not exist for the new type is reported, not silently dropped', async () => {
    const s = await blank();
    await s.run([insertOp(s, { style: { axes: { value: { title: 'Amount' } } } })]);
    const res = await s.run([update({ chartType: 'pie', allowRebuild: true, series: [
      { name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] },
    ] })]);
    expect(res.success).toBe(true);
    expect(res.chartUpdates[0].styleNotApplied?.join(' ')).toMatch(/axes/);
  });

  it('2D to 3D and scatter switches need the flag too', async () => {
    const s = await blank();
    await s.run([insertOp(s)]);
    for (const chartType of ['column3D', 'line']) {
      const res = await s.run([update({ chartType })]);
      expect(res.success).toBe(false);
      expect(JSON.stringify(res)).toMatch(/allowRebuild/);
    }
  });

  it('updateMode is in-place for a plain data update', async () => {
    const s = await blank();
    await s.run([insertOp(s)]);
    const res = await s.run([update({ title: 'New' })]);
    expect(res.chartUpdates[0]).toMatchObject({ updateMode: 'in-place', workbook: 'untouched' });
  });
});
