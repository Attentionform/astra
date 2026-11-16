import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

const FILE = join(__dirname, '..', 'fixtures', 'word-chart-with-workbook.docx');
const renamePrefix = (xml: string, from: string, to: string): string =>
  xml.replace(new RegExp(`(</?)${from}:`, 'g'), `$1${to}:`).replace(new RegExp(`(\\s)${from}:([A-Za-z])`, 'g'), `$1${to}:$2`).replace(new RegExp(`xmlns:${from}=`, 'g'), `xmlns:${to}=`);

const variants: Array<[string, (z: AdmZip) => void]> = [
  ['canonical', () => undefined],
  ['a->b in chart', (z) => z.updateFile('word/charts/chart1.xml', Buffer.from(renamePrefix(z.readAsText('word/charts/chart1.xml'), 'a', 'b')))],
  ['c->k in chart', (z) => z.updateFile('word/charts/chart1.xml', Buffer.from(renamePrefix(z.readAsText('word/charts/chart1.xml'), 'c', 'k')))],
  ['wp->zz, w->x in document', (z) => z.updateFile('word/document.xml', Buffer.from(renamePrefix(renamePrefix(z.readAsText('word/document.xml'), 'wp', 'zz'), 'w', 'x')))],
  ['all renamed', (z) => {
    z.updateFile('word/charts/chart1.xml', Buffer.from(renamePrefix(renamePrefix(z.readAsText('word/charts/chart1.xml'), 'a', 'b'), 'c', 'k')));
    z.updateFile('word/document.xml', Buffer.from(renamePrefix(renamePrefix(z.readAsText('word/document.xml'), 'wp', 'zz'), 'w', 'x')));
  }],
];
const ops: Array<[string, (id: string) => Record<string, unknown>]> = [
  ['title', (id) => ({ action: 'chart', operation: 'update', target: id, title: 'New title' })],
  ['data', (id) => ({ action: 'chart', operation: 'update', target: id, series: [{ name: 'A', values: [5, 6, 7] }, { name: 'B', values: [1, 2, 3] }, { name: 'C', values: [4, 4, 4] }] })],
  ['type', (id) => ({ action: 'chart', operation: 'update', target: id, chartType: 'line', allowRebuild: true })],
  ['size', (id) => ({ action: 'chart', operation: 'update', target: id, widthIn: 5, heightIn: 3 })],
  ['style', (id) => ({ action: 'chart', operation: 'update', target: id, style: { dataLabels: { show: true } } })],
  ['chartFont', (id) => ({ action: 'chart', operation: 'update', target: id, style: { chartFont: { sizePt: 9 } } })],
];
/** Outcome of each operation on one variant: ok, or the refusal's first sentence with paths and counts kept. */
async function outcomes(buf: Buffer, operations: Array<[string, (id: string) => Record<string, unknown>]>): Promise<string[]> {
  const row: string[] = [];
  for (const [oname, make] of operations) {
    const parser = new DocxParser(buf);
    const ast: any = await parser.getAST();
    const node = ast.sections.flatMap((s: any) => s.nodes).find((n: any) => n.type === 'chart');
    if (!node) { row.push(`${oname}:NO-CHART-NODE`); continue; }
    const engine = new OperationEngine(ast, parser);
    const res: any = await engine.executeBatch({ operations: [make(node.id)] } as any);
    row.push(`${oname}:${res.success ? 'ok' : String(res.errors?.[0]?.error ?? '').split('\n')[0]}`);
  }
  return row;
}

it('imported Word chart: every operation behaves the same under any prefix variant (the contract)', async () => {
  const baseline = await outcomes(readFileSync(FILE), ops);
  expect(baseline.filter((r) => r.endsWith(':ok')).map((r) => r.split(':')[0])).toEqual(['title', 'type', 'size', 'style', 'chartFont']);
  expect(baseline.find((r) => r.startsWith('data:'))).toMatch(/not the workbook ASTra generated/); // foreign workbook policy
  for (const [vname, mutate] of variants.slice(1)) {
    const z = new AdmZip(readFileSync(FILE)); mutate(z);
    expect(await outcomes(z.toBuffer(), ops), vname).toEqual(baseline);
  }
}, 300000);

it('ASTra-made chart: data, title, type, size and style all work under every prefix variant', async () => {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const made: any = await engine.executeBatch({ operations: [{
    action: 'chart', operation: 'insert', anchor: ast.sections[0].nodes[0].id, id: 'c1', chartType: 'column', title: 'T',
    categories: ['a', 'b', 'c'], series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }, { name: 'C', values: [2, 2, 2] }],
  }] } as any);
  expect(made.success, JSON.stringify(made.errors)).toBe(true);
  const generated = parser.rezip(ast);
  const chartPart = new AdmZip(generated).getEntries().map((e) => e.entryName).find((n) => /^word\/charts\/chart\d+\.xml$/.test(n))!;
  const gen: Array<[string, (z: AdmZip) => void]> = [
    ['canonical', () => undefined],
    ['all renamed', (z) => {
      z.updateFile(chartPart, Buffer.from(renamePrefix(renamePrefix(z.readAsText(chartPart), 'a', 'b'), 'c', 'k')));
      z.updateFile('word/document.xml', Buffer.from(renamePrefix(renamePrefix(z.readAsText('word/document.xml'), 'wp', 'zz'), 'w', 'x')));
    }],
  ];
  const all: Array<[string, (id: string) => Record<string, unknown>]> = [
    ...ops.filter(([n]) => n !== 'data'),
    ['data', (id) => ({ action: 'chart', operation: 'update', target: id, series: [{ name: 'A', values: [9, 9, 9] }, { name: 'B', values: [1, 1, 1] }, { name: 'C', values: [5, 5, 5] }] })],
  ];
  const rows: string[][] = [];
  for (const [, mutate] of gen) { const z = new AdmZip(generated); mutate(z); rows.push(await outcomes(z.toBuffer(), all)); }
  expect(rows[0].every((r) => r.endsWith(':ok')), rows[0].join(' | ')).toBe(true);
  expect(rows[1]).toEqual(rows[0]);
}, 300000);
