/**
 * 0.5.3: chart size in inches and EMU, defaults, guide defaults, limits, alt text and layout warnings.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, parseStyleGuide, validateDocxBuffer } from '@attentionform/astra-engine';
import { composeDocument } from '../../packages/engine/src/compose/markdown';
import {
  CHART_MAX_EMU, CHART_MIN_EMU, DEFAULT_CHART_HEIGHT_2D_EMU, DEFAULT_CHART_HEIGHT_3D_EMU, DEFAULT_CHART_WIDTH_EMU,
  parseChartSize, usableAreaOf, chartLayoutWarning,
} from '../../packages/engine/src/operations/chart-size';

const IN = 914_400;

async function session(guide?: unknown) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser, guide ? { styleGuide: guide as any } : undefined);
  const zip = () => new AdmZip(parser.rezip(ast));
  const docXml = () => zip().readAsText('word/document.xml');
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  const anchor = ast.sections[0].nodes[0].id as string;
  return { run, docXml, zip, anchor, parser, ast };
}

async function fails(s: { run: (o: unknown[]) => Promise<any> }, ops: unknown[], re: RegExp) {
  const res = await s.run(ops);
  expect(res.success).toBe(false);
  expect(JSON.stringify(res)).toMatch(re);
}

const chartOp = (anchor: string, extra: Record<string, unknown> = {}) => ({
  action: 'chart', operation: 'insert', anchor, id: 'c1', chartType: 'column', title: 'Revenue',
  categories: ['Q1', 'Q2'], series: [{ name: 'A', values: [1, 2] }], ...extra,
});

function extentOf(xml: string, index = 0): { cx: number; cy: number } {
  const all = [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)];
  return { cx: Number(all[index][1]), cy: Number(all[index][2]) };
}
const docPrOf = (xml: string, index = 0): string => [...xml.matchAll(/<wp:docPr [^>]*\/>/g)][index][0];

describe('chart size on insert', () => {
  it('2D default is 6 in x 2.6 in, 3D default is 6 in x 3.5 in', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor), chartOp(s.anchor, { id: 'c2', chartType: 'bar3D' })]);
    const xml = s.docXml();
    expect(extentOf(xml, 0)).toEqual({ cx: DEFAULT_CHART_WIDTH_EMU, cy: DEFAULT_CHART_HEIGHT_2D_EMU });
    expect(extentOf(xml, 1)).toEqual({ cx: DEFAULT_CHART_WIDTH_EMU, cy: DEFAULT_CHART_HEIGHT_3D_EMU });
    expect(DEFAULT_CHART_HEIGHT_2D_EMU).toBe(2_377_440);
  });

  it('widthIn and heightIn become exact EMU', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor, { widthIn: 5, heightIn: 2 })]);
    expect(extentOf(s.docXml())).toEqual({ cx: 5 * IN, cy: 2 * IN });
  });

  it('EMU still works, and a dimension that is not given falls back on its own', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor, { widthEmu: 4_000_000 })]);
    expect(extentOf(s.docXml())).toEqual({ cx: 4_000_000, cy: DEFAULT_CHART_HEIGHT_2D_EMU });
    const t = await session();
    await t.run([chartOp(t.anchor, { heightIn: 3 })]);
    expect(extentOf(t.docXml())).toEqual({ cx: DEFAULT_CHART_WIDTH_EMU, cy: 3 * IN });
  });

  it('giving the Emu and the In form of one dimension is an error', async () => {
    const s = await session();
    await fails(s, [chartOp(s.anchor, { widthIn: 5, widthEmu: 4_000_000 })], /width is given twice/);
    await fails(s, [chartOp(s.anchor, { heightIn: 2, heightEmu: 2_000_000 })], /height is given twice/);
  });

  it('applies the limits 0.1 in to 22 in inclusive, after rounding to EMU', () => {
    expect(parseChartSize({ widthIn: 0.1 })).toEqual({ widthEmu: CHART_MIN_EMU });
    expect(parseChartSize({ widthIn: 22 })).toEqual({ widthEmu: CHART_MAX_EMU });
    expect(parseChartSize({ widthIn: 22.0000001 })).toEqual({ widthEmu: CHART_MAX_EMU });
    expect(parseChartSize({ heightEmu: CHART_MIN_EMU })).toEqual({ heightEmu: CHART_MIN_EMU });
    for (const bad of [0.099, 0, -1, 22.01, 100, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(() => parseChartSize({ widthIn: bad })).toThrow(/outside the supported range|finite number/);
    }
    expect(() => parseChartSize({ heightEmu: CHART_MAX_EMU + 1 })).toThrow(/0.1 to 22 inches/);
    expect(() => parseChartSize({ widthIn: '5' as never })).toThrow(/finite number/);
  });

  it('an out-of-range size on insert is refused and writes nothing', async () => {
    const s = await session();
    await fails(s, [chartOp(s.anchor, { widthIn: 30 })], /outside the supported range/);
    expect(s.docXml()).not.toContain('<c:chart');
    expect(s.zip().getEntry('word/charts/chart1.xml')).toBeNull();
  });

  it('every size written is valid for Word and the validator', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor, { widthIn: 0.1, heightIn: 0.1 }), chartOp(s.anchor, { id: 'c2', widthIn: 22, heightIn: 22 })]);
    const docx = await validateDocxBuffer(s.parser.rezip());
    expect(docx.errors).toEqual([]);
  });
});

describe('chart size from the style guide', () => {
  const guideText = (lines: string) => parseStyleGuide(`# Style Guide: t\n\n## Charts\n${lines}`);

  it('uses the guide defaults, 2D and 3D separately, and the operation wins per dimension', async () => {
    const guide = guideText('- Width: 5 in\n- Height: 2 in\n- 3D Width: 4 in\n- 3D Height: 3 in\n').config;
    const s = await session(guide);
    await s.run([chartOp(s.anchor), chartOp(s.anchor, { id: 'c2', chartType: 'pie3D' }), chartOp(s.anchor, { id: 'c3', heightIn: 4 })]);
    const xml = s.docXml();
    expect(extentOf(xml, 0)).toEqual({ cx: 5 * IN, cy: 2 * IN });
    expect(extentOf(xml, 1)).toEqual({ cx: 4 * IN, cy: 3 * IN });
    expect(extentOf(xml, 2)).toEqual({ cx: 5 * IN, cy: 4 * IN });
  });

  it('reads cm, mm and pt, and a number with no unit as inches', () => {
    const cfg = guideText('- Width: 15.24 cm\n- Height: 50.8 mm\n- 3D Width: 432 pt\n- 3D Height: 3\n').config.charts!;
    expect(cfg.widthIn).toBeCloseTo(6, 3);
    expect(cfg.heightIn).toBeCloseTo(2, 3);
    expect(cfg.width3dIn).toBeCloseTo(6, 3);
    expect(cfg.height3dIn).toBe(3);
  });

  it('never reinterprets an unknown unit: the value is not used and the warning names it', () => {
    const parsed = guideText('- Width: 6 furlong\n- Height: 12 px\n');
    expect(parsed.config.charts?.widthIn).toBeUndefined();
    expect(parsed.config.charts?.heightIn).toBeUndefined();
    expect(parsed.warnings.join('\n')).toContain("unit 'furlong' is not supported");
    expect(parsed.warnings.join('\n')).toContain("unit 'px' is not supported");
  });

  it('text sizes accept pt only; cm on a text size is not read as points', () => {
    const parsed = guideText('- Font Size: 9 pt\n- Title Size: 1 cm\n');
    expect(parsed.config.charts?.fontSizePt).toBe(9);
    expect(parsed.config.charts?.titleSizePt).toBeUndefined();
    expect(parsed.warnings.join('\n')).toContain("unit 'cm' is not supported (use pt)");
  });

  it('values outside the range or not numbers fall back to the built-in default', async () => {
    const parsed = guideText('- Width: 40 in\n- Height: banana\n');
    expect(parsed.config.charts?.widthIn).toBeUndefined();
    expect(parsed.config.charts?.heightIn).toBeUndefined();
    expect(parsed.warnings.length).toBeGreaterThanOrEqual(2);
    const s = await session(parsed.config);
    await s.run([chartOp(s.anchor)]);
    expect(extentOf(s.docXml())).toEqual({ cx: DEFAULT_CHART_WIDTH_EMU, cy: DEFAULT_CHART_HEIGHT_2D_EMU });
  });

  it('a guide without a Charts section changes nothing and adds no warning', () => {
    const parsed = parseStyleGuide('# Style Guide: t\n\n## Tables\n- Border Color: #CBD5E1\n');
    expect(parsed.config.charts).toBeUndefined();
    expect(parsed.warnings.join('\n')).not.toMatch(/chart/i);
  });

  it('keeps the remaining keys: gridline colour, font family, sizes', () => {
    const cfg = guideText('- Gridline Color: #e2e8f0\n- Font Family: Calibri\n- Font Size: 9 pt\n- Title Size: 12\n').config.charts!;
    expect(cfg.gridlineColor).toBe('E2E8F0');
    expect(cfg.fontFamily).toBe('Calibri');
    expect(cfg.fontSizePt).toBe(9);
    expect(cfg.titleSizePt).toBe(12);
  });
});

describe('chart alt text', () => {
  it('insert writes it on the drawing, escaped, and not on the chart part', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor, { altText: 'Revenue "Q1" < Q2 & more' })]);
    const dp = docPrOf(s.docXml());
    expect(dp).toContain('descr="Revenue &quot;Q1&quot; &lt; Q2 &amp; more"');
    expect(s.zip().readAsText('word/charts/chart1.xml')).not.toContain('Q1&quot;');
    const docx = await validateDocxBuffer(s.parser.rezip());
    expect(docx.errors).toEqual([]);
  });

  it('no alt text means no descr attribute', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    expect(docPrOf(s.docXml())).not.toContain('descr=');
  });

  it('update sets, replaces and clears it, and an alt-text-only update leaves the chart part and workbook untouched', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    const chartBefore = s.zip().readAsText('word/charts/chart1.xml');
    const wbBefore = s.zip().getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData().toString('base64');
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', altText: 'First' }]);
    expect(docPrOf(s.docXml())).toContain('descr="First"');
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', altText: 'Second' }]);
    expect(docPrOf(s.docXml())).toContain('descr="Second"');
    expect(docPrOf(s.docXml()).match(/descr=/g)).toHaveLength(1);
    expect(s.zip().readAsText('word/charts/chart1.xml')).toBe(chartBefore);
    expect(s.zip().getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData().toString('base64')).toBe(wbBefore);
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', altText: '' }]);
    expect(docPrOf(s.docXml())).not.toContain('descr=');
  });

  it('rejects non-string and over-long alt text', async () => {
    const s = await session();
    await fails(s, [chartOp(s.anchor, { altText: 'x'.repeat(5001) })], /longer than 5000/);
    await s.run([chartOp(s.anchor)]);
    await fails(s, [{ action: 'chart', operation: 'update', target: 'c1', altText: 42 }], /altText/);
  });

  it('alt text goes to the right chart when a document has two', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor), chartOp(s.anchor, { id: 'c2', title: 'Second' })]);
    await s.run([{ action: 'chart', operation: 'update', target: 'c2', altText: 'Only the second' }]);
    const xml = s.docXml();
    const withAlt = [...xml.matchAll(/<wp:docPr [^>]*\/>/g)].filter((m) => m[0].includes('descr='));
    expect(withAlt).toHaveLength(1);
    expect(withAlt[0][0]).toContain('Only the second');
  });
});

describe('chart size on update', () => {
  it('changes only the dimension that is given', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor, { widthIn: 5, heightIn: 3 })]);
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', widthIn: 4 }]);
    expect(extentOf(s.docXml())).toEqual({ cx: 4 * IN, cy: 3 * IN });
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', heightEmu: 2_000_000 }]);
    expect(extentOf(s.docXml())).toEqual({ cx: 4 * IN, cy: 2_000_000 });
  });

  it('a size-only update does not touch the chart part or the workbook', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    const chartBefore = s.zip().readAsText('word/charts/chart1.xml');
    const wbBefore = s.zip().getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData().toString('base64');
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', widthIn: 4, heightIn: 2 }]);
    expect(s.zip().readAsText('word/charts/chart1.xml')).toBe(chartBefore);
    expect(s.zip().getEntry('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')!.getData().toString('base64')).toBe(wbBefore);
  });

  it('resizes the second of two charts only', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor), chartOp(s.anchor, { id: 'c2', title: 'Second' })]);
    await s.run([{ action: 'chart', operation: 'update', target: 'c2', widthIn: 3, heightIn: 2 }]);
    const xml = s.docXml();
    expect(extentOf(xml, 0)).toEqual({ cx: DEFAULT_CHART_WIDTH_EMU, cy: DEFAULT_CHART_HEIGHT_2D_EMU });
    expect(extentOf(xml, 1)).toEqual({ cx: 3 * IN, cy: 2 * IN });
  });

  it('size together with a title change applies both', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', title: 'Renamed', widthIn: 4 }]);
    expect(extentOf(s.docXml()).cx).toBe(4 * IN);
    expect(s.zip().readAsText('word/charts/chart1.xml')).toContain('<a:t>Renamed</a:t>');
  });

  it('works on a floating chart', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor, { style: { floating: { horizontalAlign: 'right', wrap: 'square' } } })]);
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', widthIn: 3, heightIn: 2 }]);
    const xml = s.docXml();
    expect(xml).toContain('<wp:anchor');
    expect(extentOf(xml)).toEqual({ cx: 3 * IN, cy: 2 * IN });
  });

  it('an out-of-range size is refused and the file is unchanged; the call also needs a field', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    const before = s.docXml();
    await fails(s, [{ action: 'chart', operation: 'update', target: 'c1', widthIn: 25 }], /outside the supported range/);
    expect(s.docXml()).toBe(before);
    await fails(s, [{ action: 'chart', operation: 'update', target: 'c1' }], /no mutable fields/);
  });

  it('a size update survives save and reopen, and the file is valid', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    await s.run([{ action: 'chart', operation: 'update', target: 'c1', widthIn: 4, heightIn: 2, altText: 'Kept' }]);
    const buf = s.parser.rezip(s.ast);
    const reopened = new AdmZip(buf).readAsText('word/document.xml');
    expect(extentOf(reopened)).toEqual({ cx: 4 * IN, cy: 2 * IN });
    expect(docPrOf(reopened)).toContain('descr="Kept"');
    const docx = await validateDocxBuffer(buf);
    expect(docx.errors).toEqual([]);
  });
});

describe('layout warnings', () => {
  it('names a chart wider than the text width or taller than the usable page height', () => {
    const area = usableAreaOf({ orientation: 'portrait', paperSize: 'Letter' });
    expect(area.widthEmu).toBe(Math.round(6.5 * IN));
    expect(area.heightEmu).toBe(9 * IN);
    expect(chartLayoutWarning({ widthEmu: 6 * IN, heightEmu: 2.6 * IN }, area, false)).toBeUndefined();
    expect(chartLayoutWarning({ widthEmu: 8 * IN, heightEmu: 2 * IN }, area, false)).toMatch(/layoutWarning: .*8 in wide, the text width is 6.5 in/);
    expect(chartLayoutWarning({ widthEmu: 5 * IN, heightEmu: 10 * IN }, area, false)).toMatch(/10 in tall, the usable page height is 9 in/);
    expect(chartLayoutWarning({ widthEmu: 8 * IN, heightEmu: 10 * IN }, area, false)).toMatch(/wide.*and.*tall/);
    expect(chartLayoutWarning({ widthEmu: 8 * IN, heightEmu: 10 * IN }, area, true)).toBeUndefined();
  });

  it('A4 and landscape change the area', () => {
    const a4 = usableAreaOf({ orientation: 'portrait', paperSize: 'A4' });
    expect(a4.widthEmu).toBeLessThan(Math.round(6.5 * IN));
    expect(a4.widthEmu).toBeGreaterThan(6 * IN);
    const land = usableAreaOf({ orientation: 'landscape', paperSize: 'Letter' });
    expect(land.widthEmu).toBe(9 * IN);
    expect(land.heightEmu).toBe(Math.round(6.5 * IN));
  });

  it('a wide inline chart is written and reported as a warning in the batch result; a floating one is not', async () => {
    const s = await session();
    const wide = await s.run([chartOp(s.anchor, { widthIn: 8 })]);
    expect(extentOf(s.docXml()).cx).toBe(8 * IN);
    expect((wide as any).warnings.join('\n')).toMatch(/layoutWarning/);
    const t = await session();
    const floating = await t.run([chartOp(t.anchor, { widthIn: 8, style: { floating: { wrap: 'square' } } })]);
    expect((floating as any).warnings.join('\n')).not.toMatch(/layoutWarning/);
  });

  it('resizing to a wide chart on update also warns', async () => {
    const s = await session();
    await s.run([chartOp(s.anchor)]);
    const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', widthIn: 9 }]);
    expect((res as any).warnings.join('\n')).toMatch(/layoutWarning.*9 in wide/);
  });
});

describe('Markdown chart JSON', () => {
  async function compose(chartJson: unknown, markdownAlt = 'Alt from markdown') {
    const fs = await import('node:fs');
    const os = await import('node:os');
    const path = await import('node:path');
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-chart-'));
    fs.writeFileSync(path.join(dir, 'c.json'), JSON.stringify(chartJson));
    const parser = DocxParser.createNewDocument('blank');
    await composeDocument(parser, `# T\n\n![${markdownAlt}](c.json)\n`, { allowedDir: dir });
    return new AdmZip(parser.rezip());
  }
  const base = { chartType: 'column', title: 'Revenue', categories: ['Q1', 'Q2'], series: [{ name: 'A', values: [1, 2] }] };

  it('reads size, alt text and style from the chart JSON', async () => {
    const zip = await compose({ ...base, widthIn: 5, heightIn: 2, altText: 'From the JSON', style: { legend: { position: 't' } } });
    const xml = zip.readAsText('word/document.xml');
    expect(extentOf(xml)).toEqual({ cx: 5 * IN, cy: 2 * IN });
    expect(docPrOf(xml)).toContain('descr="From the JSON"');
    expect(zip.readAsText('word/charts/chart1.xml')).toContain('<c:legendPos val="t"/>');
  });

  it('the Markdown image text is the alt text when the JSON has none, and the JSON wins when it has one', async () => {
    const withoutJsonAlt = await compose(base, 'Alt from markdown');
    expect(docPrOf(withoutJsonAlt.readAsText('word/document.xml'))).toContain('descr="Alt from markdown"');
    const withJsonAlt = await compose({ ...base, altText: 'JSON wins' }, 'Alt from markdown');
    expect(docPrOf(withJsonAlt.readAsText('word/document.xml'))).toContain('descr="JSON wins"');
  });

  it('with neither size nor alt text the chart gets the defaults', async () => {
    const zip = await compose(base, '');
    const xml = zip.readAsText('word/document.xml');
    expect(extentOf(xml)).toEqual({ cx: DEFAULT_CHART_WIDTH_EMU, cy: DEFAULT_CHART_HEIGHT_2D_EMU });
  });

  it('refuses a bad size with the line and a remedy', async () => {
    await expect(compose({ ...base, widthIn: 'wide' })).rejects.toThrow(/widthIn must be a number/);
    await expect(compose({ ...base, widthIn: 40 })).rejects.toThrow(/outside the supported range/);
    await expect(compose({ ...base, widthIn: 5, widthEmu: 4_000_000 })).rejects.toThrow(/width is given twice/);
    await expect(compose({ ...base, style: 'bold' })).rejects.toThrow(/style must be an object/);
  });
});
