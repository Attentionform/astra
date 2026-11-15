/**
 * 0.5.3: charts saved by Word carry extension content: `mc:AlternateContent` with a c14 style, `c:extLst` at the
 * chart-space, axis, series and data-label levels, c15 elements whose local names (`spPr`, `layout`) match chart
 * elements, and c16 unique ids. Style and update operations must keep all of it, keep every `extLst` last in its
 * parent, and never patch an element that only looks like a chart element.
 */
import { describe, expect, it } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART = 'word/charts/chart1.xml';
const C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart';
const C16 = 'http://schemas.microsoft.com/office/drawing/2014/chart';
const C14 = 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const uniqueId = (n: number) => `<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="${C16}"><c16:uniqueId val="{0000000${n}-AAAA-4BBB-8CCC-000000000001}"/></c:ext></c:extLst>`;
const dlblExt = `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="${C15}"><c15:spPr xmlns:c15="${C15}"><a:prstGeom prst="wedgeRectCallout"><a:avLst/></a:prstGeom></c15:spPr></c:ext></c:extLst>`;
const MARKERS = ['c14:style val="102"', 'c15:spPr', 'c16:uniqueId val="{00000001', 'c16:uniqueId val="{00000002', 'c16r3:dispNaAsBlank', 'wedgeRectCallout', 'urn:ext-axis-marker'];

/** Word-style extension content added to a chart the engine generated. */
function withExtensions(xml: string): string {
  let out = xml;
  out = out.replace('<c:lang val="en-US"/>', `<c:lang val="en-US"/><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="c14" xmlns:c14="${C14}"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>`);
  let n = 0;
  out = out.replace(/(<c:cat>)/g, `<c:dLbls><c:spPr><a:noFill/></c:spPr><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>${dlblExt}</c:dLbls>$1`);
  out = out.replace(/<\/c:ser>/g, () => `${uniqueId(++n)}</c:ser>`);
  out = out.replace(/<\/c:catAx>/, '<c:extLst><c:ext uri="{urn:ext-axis-marker}" xmlns:c16="' + C16 + '"><c16:axisMarker val="1"/></c:ext></c:extLst></c:catAx>');
  out = out.replace('</c:chartSpace>', `<c:extLst><c:ext uri="{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst></c:chartSpace>`);
  return out;
}

async function session() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const zip = parser.getZipAdapter();
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  await run([{
    action: 'chart', operation: 'insert', anchor: ast.sections[0].nodes[0].id, id: 'c1', chartType: 'column', title: 'Revenue',
    categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'A', values: [1, 2, 3], color: '0066CC' }, { name: 'B', values: [3, 2, 1], color: '10B981' }],
  }]);
  const injected = withExtensions(zip.readText(CHART)).replace('xmlns:r=', `xmlns:mc="${MC}" xmlns:r=`);
  // keep the namespace declarations the markers need on the root so the part stays one well-formed document
  zip.writeText(CHART, injected);
  return { parser, ast, run, zip, xml: () => zip.readText(CHART) };
}

/** Every `extLst` is the last element child of its parent; the part parses. */
function expectSound(xml: string): void {
  const doc = new DOMParser({ errorHandler: { warning: () => undefined, error: (e) => { throw new Error(String(e)); }, fatalError: (e) => { throw new Error(String(e)); } } })
    .parseFromString(xml, 'text/xml') as unknown as Document;
  for (const ext of Array.from(doc.getElementsByTagNameNS(C_NS, 'extLst'))) {
    let next = ext.nextSibling;
    while (next && next.nodeType !== 1) next = next.nextSibling;
    expect(next, `extLst must be last in <${(ext.parentNode as Element).tagName}>`).toBeNull();
  }
  const chart = doc.getElementsByTagNameNS(C_NS, 'chart')[0];
  const direct = (name: string): number => Array.from(chart.childNodes).filter((n) => n.nodeType === 1 && (n as Element).localName === name).length;
  expect(direct('title')).toBeLessThanOrEqual(1);
  expect(direct('legend')).toBeLessThanOrEqual(1);
  expect(direct('plotArea')).toBe(1);
}
const keptMarkers = (xml: string): string[] => MARKERS.filter((m) => xml.includes(m));

describe('extension content survives style and update operations', () => {
  it('the fixture itself is sound and carries every marker', async () => {
    const s = await session();
    expectSound(s.xml());
    expect(keptMarkers(s.xml())).toEqual(MARKERS);
  });

  const styleCases: Array<[string, Record<string, unknown>]> = [
    ['legend', { legend: { position: 'b' } }],
    ['title format', { titleFormat: { sizePt: 20, bold: true, color: '1F3864' } }],
    ['data labels on', { dataLabels: { show: true, numberFormat: '0.0' } }],
    ['data labels off', { dataLabels: { show: false } }],
    ['series colours', { series: [{ seriesNumber: 1, points: [{ pointNumber: 2, color: 'FF0000' }] }, { seriesNumber: 2, points: [{ pointNumber: 1, color: '00FF00' }] }] }],
    ['axis titles and gridlines', { axes: { category: { title: 'Quarter', labelSizePt: 9 }, value: { title: 'Amount', gridlineColor: 'CCCCCC', showAxisLine: true } } }],
    ['chart font', { chartFont: { family: 'Calibri', sizePt: 11, color: '404040' } }],
    ['bars', { bars: { gapPercent: 80, overlapPercent: 10 } }],
    ['chart and plot area', { chartArea: { fillColor: 'F2F2F2' }, plotArea: { fillColor: 'FFFFFF' } }],
    ['trendline', { trendline: { type: 'linear', series: undefined, seriesNumbers: [1] } }],
  ];
  for (const [label, style] of styleCases) {
    it(`style: ${label} keeps the extensions and the part sound`, async () => {
      const s = await session();
      const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', style }]);
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      expectSound(s.xml());
      expect(keptMarkers(s.xml())).toEqual(MARKERS);
    });
  }

  it('all style keys together, then a second call, keep the extensions and one legend', async () => {
    const s = await session();
    const style = {
      legend: { position: 'r' }, titleFormat: { sizePt: 18 }, dataLabels: { show: true },
      axes: { value: { title: 'Amount', gridlineColor: 'BBBBBB' } }, chartFont: { family: 'Arial', sizePt: 10 },
      bars: { gapPercent: 60 }, chartArea: { fillColor: 'FAFAFA' },
    };
    for (let i = 0; i < 2; i++) {
      const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', style }]);
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
    }
    expectSound(s.xml());
    expect(keptMarkers(s.xml())).toEqual(MARKERS);
  });

  it('title, size and alt text keep the extensions', async () => {
    const s = await session();
    const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', title: 'New title', widthIn: 5, heightIn: 3, altText: 'Chart' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expectSound(s.xml());
    expect(keptMarkers(s.xml())).toEqual(MARKERS);
  });

  it('a data update on the same series keeps the extensions and the series-level data labels', async () => {
    const s = await session();
    const res = await s.run([{
      action: 'chart', operation: 'update', target: 'c1', categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'A', values: [9, 8, 7] }, { name: 'B', values: [1, 2, 3] }],
    }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expectSound(s.xml());
    expect(keptMarkers(s.xml())).toEqual(MARKERS);
    expect(s.xml()).toContain('<c:v>9</c:v>');
  });

  it('adding a series at the end keeps the existing series extensions and ends with a valid part', async () => {
    const s = await session();
    const res = await s.run([{
      action: 'chart', operation: 'update', target: 'c1', categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] }],
    }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expectSound(s.xml());
    expect(keptMarkers(s.xml())).toEqual(MARKERS);
  });

  it('removing a series drops that series and keeps the other one and the chart-level extensions', async () => {
    const s = await session();
    const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', removeSeries: [2] }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expectSound(s.xml());
    expect(s.xml()).toContain('c16:uniqueId val="{00000001');
    expect(s.xml()).not.toContain('c16:uniqueId val="{00000002');
    for (const m of ['c14:style val="102"', 'c16r3:dispNaAsBlank', 'urn:ext-axis-marker']) expect(s.xml()).toContain(m);
  });

  it('a type change inside the family keeps the extensions', async () => {
    const s = await session();
    const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', chartType: 'columnStacked' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expectSound(s.xml());
    expect(keptMarkers(s.xml())).toEqual(MARKERS);
  });

  it('the c15 element named like a chart element is never patched as one', async () => {
    const s = await session();
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', style: { dataLabels: { show: true, numberFormat: '0.00' }, chartArea: { fillColor: 'EEEEEE' } } }]);
    const xml = s.xml();
    for (const m of xml.matchAll(/<c15:spPr[^>]*>([\s\S]*?)<\/c15:spPr>/g)) {
      expect(m[1]).toBe('<a:prstGeom prst="wedgeRectCallout"><a:avLst/></a:prstGeom>');
    }
    expect(xml.match(/<c15:spPr/g)?.length).toBe(2);
  });

  it('the package validates after the whole run', async () => {
    const s = await session();
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', style: { legend: { position: 'b' }, chartFont: { family: 'Arial' }, dataLabels: { show: true } } }]);
    const report = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(report.errors).toEqual([]);
  });
});
