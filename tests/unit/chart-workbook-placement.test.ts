/**
 * A private or replacement workbook must be a well-formed package part: found from the chart that uses it (also for
 * a chart outside word/charts/), typed as an XLSX part, and an XLSX part in format and in declared type.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { applyWorkbookDecision, planWorkbook, relativeTarget } from '../../packages/engine/src/operations/chart-workbook';

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  return { parser, ast, engine, zip: parser.getZipAdapter(), run: (operations: unknown[]) => engine.executeBatch({ operations } as any) };
}
type Session = Awaited<ReturnType<typeof open>>;

const insertOp = (anchor: string, id: string) => ({
  action: 'chart', operation: 'insert', anchor, id, chartType: 'column', title: id,
  categories: ['Q1', 'Q2', 'Q3'], series: [{ name: id, values: [1, 2, 3] }],
});

/** The node id of the chart whose first series is called `name` (ids from an insert do not survive a reopen). */
function chartNamed(s: Session, name: string): string {
  const nodes = s.ast.sections.flatMap((sec: any) => sec.nodes).filter((n: any) => n.type === 'chart');
  const node = nodes.find((n: any) => n.metadata?.series?.[0]?.name === name);
  if (!node) throw new Error(`no chart with series ${name}`);
  return node.id as string;
}

/** Two charts; chart2 is made to share chart1's workbook. Returns the package bytes. */
async function sharedPackage(): Promise<{ buf: Buffer; workbook: string }> {
  const s = await open(createBlankDocxBuffer('blank'));
  await s.run([insertOp(s.ast.sections[0].nodes[0].id, 'c1'), insertOp(s.ast.sections[0].nodes[0].id, 'c2')]);
  const books = s.zip.entryNames!().filter((n) => /^word\/embeddings\/.*\.xlsx$/.test(n)).sort();
  const rels2 = 'word/charts/_rels/chart2.xml.rels';
  s.zip.writeText(rels2, s.zip.readText(rels2).replace(/Target="[^"]*embeddings\/[^"]*"/, `Target="../embeddings/${books[0].split('/').pop()}"`));
  return { buf: s.parser.rezip(s.ast), workbook: books[0] };
}

/** The target part a relationship file points at (first relationship of the given type suffix). */
function targetOf(zip: AdmZip, relsPath: string, from: string): string | undefined {
  const m = /<Relationship\b[^>]*\sType="[^"]*\/package"[^>]*>/.exec(zip.readAsText(relsPath));
  const target = m && /\sTarget="([^"]*)"/.exec(m[0])?.[1];
  if (!target) return undefined;
  const parts = `${from}${target}`.split('/');
  const out: string[] = [];
  for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.' && p !== '') out.push(p); }
  return out.join('/');
}
function declaredType(zip: AdmZip, part: string): string | undefined {
  const ct = zip.readAsText('[Content_Types].xml');
  for (const m of ct.matchAll(/<Override\b[^>]*>/g)) {
    if (new RegExp(`\\sPartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i').test(m[0])) return /\sContentType="([^"]*)"/.exec(m[0])?.[1];
  }
  const ext = part.split('.').pop()!.toLowerCase();
  for (const m of ct.matchAll(/<Default\b[^>]*>/g)) {
    if (new RegExp(`\\sExtension="${ext}"`, 'i').test(m[0])) return /\sContentType="([^"]*)"/.exec(m[0])?.[1];
  }
  return undefined;
}

describe('private workbook for a chart outside word/charts', () => {
  // The engine edits charts stored as word/charts/chartN.xml only (a chart elsewhere is refused at update), so this
  // checks the placement code directly: the relationship target must be relative to the chart that owns it.
  it('the new relationship target resolves from the chart directory, wherever the chart is', () => {
    const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const files = new Map<string, string | Buffer>([
      ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xlsx" ContentType="${XLSX_TYPE}"/></Types>`],
      ['word/charts/sub/chart.xml', '<c:chartSpace xmlns:c="x"/>'],
      ['word/charts/sub/_rels/chart.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/package" Target="../../embeddings/Microsoft_Excel_Worksheet1.xlsx"/></Relationships>`],
      ['word/charts/other.xml', '<c:chartSpace xmlns:c="x"/>'],
      ['word/charts/_rels/other.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/package" Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/></Relationships>`],
      ['word/embeddings/Microsoft_Excel_Worksheet1.xlsx', Buffer.from('xlsx')],
    ]);
    const zip = {
      hasFile: (n: string) => files.has(n),
      readText: (n: string) => String(files.get(n)),
      readTextOrNull: (n: string) => (files.has(n) ? String(files.get(n)) : null),
      writeText: (n: string, c: string) => void files.set(n, c),
      writeBuffer: (n: string, c: Buffer) => void files.set(n, c),
      readBuffer: (n: string) => (files.has(n) ? Buffer.from(files.get(n) as string | Buffer) : null),
      entryNames: () => [...files.keys()],
      removeFile: (n: string) => void files.delete(n),
    };
    const decision = planWorkbook({
      zip, chartPath: 'word/charts/sub/chart.xml', externalDataRId: 'rId1', dataChange: true, allowRebuild: true,
      stored: { categories: ['a'], series: [{ name: 'A', values: [1] }], scatter: false },
    });
    expect(decision.action).toBe('rewrite');
    applyWorkbookDecision(zip, decision, Buffer.from('new workbook'));
    const target = /Target="([^"]+)"/.exec(String(files.get('word/charts/sub/_rels/chart.xml.rels')))![1];
    const resolved = `word/charts/sub/${target}`.split('/').reduce<string[]>((out, p) => { if (p === '..') out.pop(); else out.push(p); return out; }, []).join('/');
    expect(target).toBe('../../embeddings/Microsoft_Excel_Worksheet2.xlsx');
    expect(files.has(resolved), `${resolved} exists`).toBe(true);
    expect(files.has('word/embeddings/Microsoft_Excel_Worksheet1.xlsx')).toBe(true); // the sharer keeps its workbook
  });

  it('relativeTarget handles siblings, parents and deeper charts', () => {
    expect(relativeTarget('word/charts/', 'word/embeddings/a.xlsx')).toBe('../embeddings/a.xlsx');
    expect(relativeTarget('word/charts/sub/', 'word/embeddings/a.xlsx')).toBe('../../embeddings/a.xlsx');
    expect(relativeTarget('word/', 'word/embeddings/a.xlsx')).toBe('embeddings/a.xlsx');
    expect(relativeTarget('', 'word/embeddings/a.xlsx')).toBe('word/embeddings/a.xlsx');
  });
});

describe('private workbook content type', () => {
  it('a package that declares workbooks by individual override gets an override for the copy', async () => {
    const { buf, workbook } = await sharedPackage();
    const zip = new AdmZip(buf);
    let ct = zip.readAsText('[Content_Types].xml');
    ct = ct.replace(/<Default\b[^>]*Extension="xlsx"[^>]*\/>/i, '');
    if (!ct.includes(`PartName="/${workbook}"`)) ct = ct.replace('</Types>', `<Override PartName="/${workbook}" ContentType="${XLSX_TYPE}"/></Types>`);
    zip.updateFile('[Content_Types].xml', Buffer.from(ct));
    expect(declaredType(zip, workbook)).toBe(XLSX_TYPE);
    const s = await open(zip.toBuffer());
    const res = await s.run([{ action: 'chart', operation: 'update', target: chartNamed(s, 'c2'), allowRebuild: true, categories: ['a', 'b', 'c'] }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const out = new AdmZip(s.parser.rezip(s.ast));
    const copy = targetOf(out, 'word/charts/_rels/chart2.xml.rels', 'word/charts/');
    expect(copy).toBeDefined();
    expect(copy).not.toBe(workbook);
    expect(declaredType(out, copy!)).toBe(XLSX_TYPE);
  });
});

describe('replacing a workbook that is not an XLSX part', () => {
  it('the chart ends up pointing at an XLSX part that is declared as one', async () => {
    const s0 = await open(createBlankDocxBuffer('blank'));
    await s0.run([insertOp(s0.ast.sections[0].nodes[0].id, 'c1')]);
    const buf = s0.parser.rezip(s0.ast);
    const zip = new AdmZip(buf);
    const book = zip.getEntries().find((e) => /^word\/embeddings\/.*\.xlsx$/.test(e.entryName))!;
    const bytes = book.getData();
    // The same workbook stored as a legacy .xls part with its own default type.
    zip.deleteFile(book.entryName);
    zip.addFile('word/embeddings/Legacy.xls', Buffer.from('not an xlsx workbook'));
    zip.updateFile('word/charts/_rels/chart1.xml.rels', Buffer.from(zip.readAsText('word/charts/_rels/chart1.xml.rels').replace(/Target="[^"]*embeddings\/[^"]*"/, 'Target="../embeddings/Legacy.xls"')));
    let ct = zip.readAsText('[Content_Types].xml').replace(/<Override\b[^>]*embeddings\/[^>]*\/>/i, '');
    ct = ct.replace('</Types>', '<Default Extension="xls" ContentType="application/vnd.ms-excel"/></Types>');
    zip.updateFile('[Content_Types].xml', Buffer.from(ct));
    expect(bytes.length).toBeGreaterThan(0);
    const s = await open(zip.toBuffer());
    const res = await s.run([{ action: 'chart', operation: 'update', target: chartNamed(s, 'c1'), allowRebuild: true, categories: ['a', 'b', 'c'] }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const out = new AdmZip(s.parser.rezip(s.ast));
    const target = targetOf(out, 'word/charts/_rels/chart1.xml.rels', 'word/charts/');
    expect(target).toBeDefined();
    expect(out.getEntry(target!)).not.toBeNull();
    expect(declaredType(out, target!)).toBe(XLSX_TYPE);
    expect(() => new AdmZip(out.getEntry(target!)!.getData())).not.toThrow();
    expect(new AdmZip(out.getEntry(target!)!.getData()).getEntry('xl/worksheets/sheet1.xml')).not.toBeNull();
  });
});
