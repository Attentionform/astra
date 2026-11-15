/**
 * 0.5.3 (DN1): workbook policy for chart updates.
 *   Rule 1: an update with no data change leaves the workbook bytes identical.
 *   Rule 2: a data change rewrites only a workbook ASTra generated, unless allowRebuild.
 *   Shared workbooks get a private copy on allowRebuild; linked workbooks are never rewritten.
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
  const run = (operations: unknown[]) => engine.executeBatch({ operations } as any);
  return { parser, ast, engine, zip, run };
}
type Session = Awaited<ReturnType<typeof open>>;

const blank = () => open(createBlankDocxBuffer('blank'));
const anchorOf = (s: Session) => s.ast.sections[0].nodes[0].id as string;
const insertOp = (anchor: string, id: string, extra: Record<string, unknown> = {}) => ({
  action: 'chart', operation: 'insert', anchor, id, chartType: 'column', title: 'Revenue',
  categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }], ...extra,
});
const wb = (s: Session, path: string): Buffer => s.zip.readBuffer!(path) as Buffer;
const workbooks = (s: Session): string[] => s.zip.entryNames!().filter((n) => /^word\/embeddings\/.*\.xlsx$/.test(n)).sort();
const wordChartId = (s: Session): string => {
  const node = s.ast.sections.flatMap((sec: any) => sec.nodes).find((n: any) => n.type === 'chart');
  return node.id as string;
};
async function refused(s: Session, ops: unknown[], re: RegExp) {
  const res = await s.run(ops);
  expect(res.success).toBe(false);
  expect(JSON.stringify(res)).toMatch(re);
}

describe('workbook policy: no data change leaves the workbook alone', () => {
  const noData: Array<[string, Record<string, unknown>]> = [
    ['title', { title: 'New title' }],
    ['style', { style: { legend: { position: 'b' } } }],
    ['chartType within the family', { chartType: 'line' }],
    ['size', { widthIn: 5, heightIn: 3 }],
    ['alt text', { altText: 'A chart' }],
  ];
  for (const [label, extra] of noData) {
    it(`${label} update keeps the workbook byte-identical (ASTra chart)`, async () => {
      const s = await blank();
      await s.run([insertOp(anchorOf(s), 'c1')]);
      const [path] = workbooks(s);
      const before = Buffer.from(wb(s, path));
      const res = await s.run([{ action: 'chart', operation: 'update', target: 'c1', ...extra }]);
      expect(res.success).toBe(true);
      expect(wb(s, path).equals(before)).toBe(true);
    });
  }

  it('a title update on a chart with an edited workbook succeeds and keeps its bytes', async () => {
    const s = await open(readFileSync(WORD_CHART));
    const [path] = workbooks(s);
    const before = Buffer.from(wb(s, path));
    const res = await s.run([{ action: 'chart', operation: 'update', target: wordChartId(s), title: 'Quarterly' }]);
    expect(res.success).toBe(true);
    expect(wb(s, path).equals(before)).toBe(true);
  });

  it('a 3D view update keeps the workbook bytes', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1', { chartType: 'column3D' })]);
    const [path] = workbooks(s);
    const before = Buffer.from(wb(s, path));
    const res = await s.run([{ action: 'chart', operation: 'updateView3D', target: 'c1', view3D: { rotX: 20 } }]);
    expect(res.success).toBe(true);
    expect(wb(s, path).equals(before)).toBe(true);
  });

  it('updateView3D on a 2D chart is refused', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1')]);
    await refused(s, [{ action: 'chart', operation: 'updateView3D', target: 'c1', view3D: { rotX: 20 } }], /needs a 3D chart/);
  });
});

describe('workbook policy: data change on a generated workbook', () => {
  it('rewrites the workbook in place and the new data is in it', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1')]);
    const [path] = workbooks(s);
    const res = await s.run([{
      action: 'chart', operation: 'update', target: 'c1',
      categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'A', values: [1, 2, 3, 4] }, { name: 'B', values: [4, 3, 2, 1] }],
    }]);
    expect(res.success).toBe(true);
    expect(workbooks(s)).toEqual([path]);
    const sheet = new AdmZip(wb(s, path)).readAsText('xl/worksheets/sheet1.xml');
    expect(sheet).toContain('Q4');
  });

  it('a second data change after the first is still recognised', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1')]);
    const a = await s.run([{ action: 'chart', operation: 'update', target: 'c1', categories: ['X', 'Y', 'Z'] }]);
    const b = await s.run([{ action: 'chart', operation: 'update', target: 'c1', series: [{ name: 'A', values: [9, 8, 7] }, { name: 'B', values: [1, 1, 1] }] }]);
    expect(a.success).toBe(true);
    expect(b.success).toBe(true);
  });
});

describe('workbook policy: workbooks ASTra did not generate', () => {
  it('Word-authored workbook: data change refused without allowRebuild, nothing changed', async () => {
    const s = await open(readFileSync(WORD_CHART));
    const [path] = workbooks(s);
    const bookBefore = Buffer.from(wb(s, path));
    const chartBefore = s.zip.readText('word/charts/chart1.xml');
    await refused(s, [{
      action: 'chart', operation: 'update', target: wordChartId(s),
      categories: ['A', 'B', 'C', 'D'],
    }], /allowRebuild/);
    expect(wb(s, path).equals(bookBefore)).toBe(true);
    expect(s.zip.readText('word/charts/chart1.xml')).toBe(chartBefore);
  });

  it('Word-authored workbook: allowRebuild replaces it and the package stays valid', async () => {
    const s = await open(readFileSync(WORD_CHART));
    const [path] = workbooks(s);
    const res = await s.run([{
      action: 'chart', operation: 'update', target: wordChartId(s), allowRebuild: true,
      categories: ['A', 'B', 'C', 'D'], series: [{ name: 'S1', values: [1, 2, 3, 4] }],
    }]);
    expect(res.success).toBe(true);
    expect(new AdmZip(wb(s, path)).readAsText('xl/worksheets/sheet1.xml')).toContain('S1');
    const report = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(JSON.stringify(report)).not.toMatch(/"severity":"error"/);
  });

  it('workbook edited after generation (cache and workbook disagree) is refused', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1')]);
    const [path] = workbooks(s);
    const book = new AdmZip(wb(s, path));
    const sheet = book.readAsText('xl/worksheets/sheet1.xml').replace('<v>2</v>', '<v>222</v>');
    book.updateFile('xl/worksheets/sheet1.xml', Buffer.from(sheet, 'utf8'));
    s.zip.writeBuffer(path, book.toBuffer());
    const edited = Buffer.from(wb(s, path));
    await refused(s, [{ action: 'chart', operation: 'update', target: 'c1', categories: ['a', 'b', 'c'] }], /not the workbook ASTra generated/);
    expect(wb(s, path).equals(edited)).toBe(true);
  });

  it('a workbook with an extra sheet part is refused', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1')]);
    const [path] = workbooks(s);
    const book = new AdmZip(wb(s, path));
    book.addFile('xl/worksheets/sheet2.xml', Buffer.from('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'));
    s.zip.writeBuffer(path, book.toBuffer());
    await refused(s, [{ action: 'chart', operation: 'update', target: 'c1', series: [{ name: 'A', values: [5, 5, 5] }, { name: 'B', values: [1, 1, 1] }] }], /allowRebuild/);
  });
});

describe('workbook policy: shared and linked workbooks', () => {
  async function sharedPair() {
    const s = await blank();
    await s.run([insertOp(s.ast.sections[0].nodes[0].id, 'c1'), insertOp(s.ast.sections[0].nodes[0].id, 'c2', { title: 'Second' })]);
    const [first, second] = workbooks(s);
    const rels2 = 'word/charts/_rels/chart2.xml.rels';
    const text = s.zip.readText(rels2).replace(/Target="[^"]*embeddings\/[^"]*"/, `Target="../embeddings/${first.split('/').pop()}"`);
    s.zip.writeText(rels2, text);
    return { s, first, second };
  }

  it('shared workbook: data change refused without allowRebuild', async () => {
    const { s, first } = await sharedPair();
    const before = Buffer.from(wb(s, first));
    await refused(s, [{ action: 'chart', operation: 'update', target: 'c2', categories: ['a', 'b', 'c'] }], /also used by/);
    expect(wb(s, first).equals(before)).toBe(true);
  });

  it('shared workbook: allowRebuild gives a private copy and the other chart keeps the original', async () => {
    const { s, first } = await sharedPair();
    const before = Buffer.from(wb(s, first));
    const res = await s.run([{ action: 'chart', operation: 'update', target: 'c2', allowRebuild: true, categories: ['a', 'b', 'c'] }]);
    expect(res.success).toBe(true);
    expect(wb(s, first).equals(before)).toBe(true);
    const rels2 = s.zip.readText('word/charts/_rels/chart2.xml.rels');
    const target = /Target="\.\.\/embeddings\/([^"]+)"/.exec(rels2)![1];
    expect(`word/embeddings/${target}`).not.toBe(first);
    expect(new AdmZip(wb(s, `word/embeddings/${target}`)).readAsText('xl/worksheets/sheet1.xml')).toContain('>a<');
    expect(s.zip.readText('word/charts/_rels/chart1.xml.rels')).toContain(first.split('/').pop()!);
  });

  it('shared workbook: a no-data update (title) is allowed and writes nothing', async () => {
    const { s, first } = await sharedPair();
    const before = Buffer.from(wb(s, first));
    const res = await s.run([{ action: 'chart', operation: 'update', target: 'c2', title: 'Renamed' }]);
    expect(res.success).toBe(true);
    expect(wb(s, first).equals(before)).toBe(true);
  });

  it('linked workbook: data change refused even with allowRebuild, title update allowed', async () => {
    const s = await blank();
    await s.run([insertOp(anchorOf(s), 'c1')]);
    const rels = 'word/charts/_rels/chart1.xml.rels';
    s.zip.writeText(rels, s.zip.readText(rels).replace(/(<Relationship [^>]*embeddings[^>]*?)\/>/, '$1 TargetMode="External"/>')
      .replace(/Target="\.\.\/embeddings\/[^"]*"/, 'Target="file:///C:/data/book.xlsx"'));
    await refused(s, [{ action: 'chart', operation: 'update', target: 'c1', allowRebuild: true, categories: ['a', 'b', 'c'] }], /linked workbook/);
    const ok = await s.run([{ action: 'chart', operation: 'update', target: 'c1', title: 'Still fine' }]);
    expect(ok.success).toBe(true);
  });

  it('a refused batch leaves earlier operations rolled back', async () => {
    const { s, first } = await sharedPair();
    const res = await s.run([
      { action: 'chart', operation: 'update', target: 'c1', title: 'Changed first' },
      { action: 'chart', operation: 'update', target: 'c2', categories: ['a', 'b', 'c'] },
    ]);
    expect(res.success).toBe(false);
    expect(s.zip.readText('word/charts/chart1.xml')).not.toContain('Changed first');
    expect(wb(s, first).length).toBeGreaterThan(0);
  });
});
