import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  executeChart,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

function chartXmlFromBuffer(buf: Buffer, name: string): string | null {
  const zip = new AdmZip(buf);
  const entry = zip.getEntry(`word/charts/${name}`);
  return entry ? entry.getData().toString('utf8') : null;
}

function chartPartNames(buf: Buffer): string[] {
  const zip = new AdmZip(buf);
  return (zip.getEntries() || [])
    .map((e: { entryName: string }) => e.entryName)
    .filter((n: string) => /^word\/charts\/chart\d+\.xml$/.test(n));
}

function sheetXmlFromBuffer(buf: Buffer, workbookName: string): string {
  const zip = new AdmZip(buf);
  const wbEntry = zip.getEntry(`word/embeddings/${workbookName}`);
  if (!wbEntry) throw new Error(`${workbookName} missing from rezipped docx`);
  const wbZip = new AdmZip(wbEntry.getData());
  return wbZip.getEntry('xl/worksheets/sheet1.xml')?.getData().toString('utf8') ?? '';
}

describe('Chart update in-place (update op)', () => {
  it('update changes series data in chart1.xml with no chart2.xml created', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      id: 'chart_upd_1',
      title: 'Before Title',
      categories: ['Q1', 'Q2'],
      series: [{ name: 'ARR', values: [10, 20], color: '0066CC' }],
    } as never, ast, parser);

    executeChart({
      action: 'chart',
      operation: 'update',
      id: 'chart_upd_1',
      title: 'After Title',
      categories: ['H1', 'H2', 'H3'],
      series: [{ name: 'ARR', values: [111, 222, 333], color: '10B981' }],
    } as never, ast, parser);

    const buf = parser.rezip();
    const parts = chartPartNames(buf);
    expect(parts).toEqual(['word/charts/chart1.xml']);
    expect(chartXmlFromBuffer(buf, 'chart2.xml')).toBeNull();

    const chartXml = chartXmlFromBuffer(buf, 'chart1.xml');
    expect(chartXml).not.toBeNull();
    expect(chartXml!).toContain('After Title');
    expect(chartXml!).not.toContain('Before Title');
    expect(chartXml!).toContain('<c:v>111</c:v>');
    expect(chartXml!).toContain('<c:v>333</c:v>');
    expect(chartXml!).toContain('<c:v>H1</c:v>');
    expect(chartXml!).toContain('<a:srgbClr val="10B981"/>');

    const sheetXml = sheetXmlFromBuffer(buf, 'Microsoft_Excel_Worksheet1.xlsx');
    expect(sheetXml).toContain('111');
    expect(sheetXml).toContain('333');
    expect(sheetXml).toContain('H1');
  });

  it('update of a missing chart fails closed with E_VALIDATION', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    try {
      executeChart({
        action: 'chart',
        operation: 'update',
        series: [{ name: 'ARR', values: [1, 2] }],
      } as never, ast, parser);
      expect.unreachable('update on a chartless document should throw');
    } catch (err: unknown) {
      const e = err as Error & { code?: string; remedy?: string };
      expect(e.code).toBe('E_VALIDATION');
      expect(e.remedy).toMatch(/insert/i);
    }
    expect(chartPartNames(parser.rezip())).toEqual([]);

    executeChart({
      action: 'chart',
      operation: 'insert',
      chartType: 'column',
      id: 'chart_real',
      title: 'Real',
      categories: ['Q1'],
      series: [{ name: 'S', values: [5] }],
    } as never, ast, parser);
    try {
      executeChart({
        action: 'chart',
        operation: 'update',
        id: 'chart_ghost',
        series: [{ name: 'S', values: [9] }],
      } as never, ast, parser);
      expect.unreachable('update of an unknown chart id should throw');
    } catch (err: unknown) {
      const e = err as Error & { code?: string; remedy?: string };
      expect(e.code).toBe('E_VALIDATION');
      expect(e.remedy).toMatch(/chart_real/);
    }
    // Failed updates leave the original part untouched and unduplicated.
    const buf = parser.rezip();
    expect(chartPartNames(buf)).toEqual(['word/charts/chart1.xml']);
    expect(chartXmlFromBuffer(buf, 'chart1.xml')).toContain('<c:v>5</c:v>');
  });

  it('save (rezip) after insert+update passes validation', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    executeChart({
      action: 'chart',
      operation: 'insert',
      chartType: 'line',
      id: 'chart_val_1',
      title: 'Trend',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [{ name: 'ARR', values: [1, 2, 3], color: '0066CC' }],
    } as never, ast, parser);
    executeChart({
      action: 'chart',
      operation: 'update',
      id: 'chart_val_1',
      series: [{ name: 'ARR', values: [7, 8, 9], color: '0066CC' }],
    } as never, ast, parser);

    const buf = parser.rezip();
    const bufValidation = await validateDocxBuffer(buf);
    expect(bufValidation.ok).toBe(true);
    expect(bufValidation.errors).toEqual([]);
    const ooxmlValidation = await parser.validateOOXML();
    expect(ooxmlValidation.ok).toBe(true);
    expect(ooxmlValidation.errors).toEqual([]);
  });
});
