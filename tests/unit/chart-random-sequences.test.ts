/**
 * 0.5.3: seeded random sequences of chart operations on several starting packages, with package invariants
 * checked after every step. A failure prints its seed, starting package and step so it can be replayed.
 *
 *   FUZZ_SEEDS=40 FUZZ_STEPS=40 npx vitest run tests/unit/chart-random-sequences   (longer run; CI)
 *
 * Invariants:
 *   - every internal relationship target and every content-type override exists in the package;
 *   - no chart part or embedded workbook is left unreferenced;
 *   - chart nodes, chart parts and chart relationships agree in number, and series counts match;
 *   - a refused batch leaves every part byte-identical;
 *   - an update changes at most the targeted chart part, and no workbook unless it changes data;
 *   - the package validates with no more errors than the starting package, and reopens with the same charts.
 */
import { readFileSync } from 'node:fs';
import { appendFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const SEEDS = Number(process.env.FUZZ_SEEDS ?? 3);
const STEPS = Number(process.env.FUZZ_STEPS ?? 16);
const VALIDATE_EVERY = Number(process.env.FUZZ_VALIDATE_EVERY ?? 6);
const WORD_CHART = join(__dirname, '..', 'fixtures', 'word-chart-with-workbook.docx');

// ---------------------------------------------------------------------------------------------- random
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
type Rand = ReturnType<typeof rng>;
const int = (r: Rand, lo: number, hi: number): number => lo + Math.floor(r() * (hi - lo + 1));
const pick = <T>(r: Rand, items: readonly T[]): T => items[Math.floor(r() * items.length)] as T;
const chance = (r: Rand, p: number): boolean => r() < p;
const hex = (r: Rand): string => Array.from({ length: 6 }, () => '0123456789ABCDEF'[int(r, 0, 15)]).join('');

// ---------------------------------------------------------------------------------------------- session
async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  return { parser, ast, run };
}
type Session = Awaited<ReturnType<typeof open>>;
const chartNodes = (s: Session): any[] => s.ast.sections.flatMap((sec: any) => sec.nodes).filter((n: any) => n.type === 'chart' || n.type === 'chart3d');
const snapshot = (s: Session): Buffer => s.parser.rezip(s.ast);

// ---------------------------------------------------------------------------------------------- package checks
const REL_TAG = /<(?:[\w.-]+:)?Relationship\b[^>]*>/g;
const OVERRIDE_TAG = /<(?:[\w.-]+:)?Override\b[^>]*>/g;
const attrOf = (tag: string, name: string): string | undefined => new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag)?.slice(1).find((v) => v !== undefined);

function resolve(sourcePart: string, target: string): string {
  const base = target.startsWith('/') ? target.slice(1) : `${sourcePart.slice(0, sourcePart.lastIndexOf('/') + 1)}${target}`;
  const out: string[] = [];
  for (const seg of base.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}
function sourceOfRels(rels: string): string {
  if (rels === '_rels/.rels') return '';
  const m = /^(.*)_rels\/(.*)\.rels$/.exec(rels);
  return m ? `${m[1]}${m[2]}` : rels;
}

interface PackageReport { problems: string[]; chartParts: string[]; chartRelCount: number; serCounts: number[] }
function inspectPackage(buffer: Buffer): PackageReport {
  const zip = new AdmZip(buffer);
  const names = new Set(zip.getEntries().map((e) => e.entryName));
  const lower = new Set([...names].map((n) => n.toLowerCase()));
  const problems: string[] = [];
  const referenced = new Set<string>();
  let chartRelCount = 0;
  for (const name of names) {
    if (!name.endsWith('.rels')) continue;
    const xml = zip.readAsText(name);
    const source = sourceOfRels(name);
    for (const tag of xml.match(REL_TAG) ?? []) {
      if ((attrOf(tag, 'TargetMode') ?? '').toLowerCase() === 'external') continue;
      const target = attrOf(tag, 'Target');
      if (target === undefined) continue;
      const where = resolve(source, target);
      referenced.add(where.toLowerCase());
      if (!lower.has(where.toLowerCase())) problems.push(`dangling relationship ${name} -> ${where}`);
      if (name === 'word/_rels/document.xml.rels' && /\/chart$/.test(attrOf(tag, 'Type') ?? '')) chartRelCount++;
    }
  }
  const types = names.has('[Content_Types].xml') ? zip.readAsText('[Content_Types].xml') : '';
  for (const tag of types.match(OVERRIDE_TAG) ?? []) {
    const part = (attrOf(tag, 'PartName') ?? '').replace(/^\//, '');
    if (part && !lower.has(part.toLowerCase())) problems.push(`content-type override without part ${part}`);
  }
  const chartParts = [...names].filter((n) => /^word\/charts\/chart\d+\.xml$/.test(n)).sort();
  for (const name of names) {
    if (/^word\/charts\/chart\d+\.xml$/.test(name) || /^word\/embeddings\//.test(name)) {
      if (!referenced.has(name.toLowerCase())) problems.push(`unreferenced part ${name}`);
    }
  }
  const serCounts = chartParts.map((p) => (zip.readAsText(p).match(/<(?:[\w.-]+:)?ser[\s>]/g) ?? []).length);
  return { problems, chartParts, chartRelCount, serCounts };
}

function digest(buffer: Buffer): Map<string, string> {
  const out = new Map<string, string>();
  for (const e of new AdmZip(buffer).getEntries()) {
    if (e.entryName.startsWith('docProps/')) continue;
    out.set(e.entryName, createHash('sha1').update(e.getData()).digest('hex'));
  }
  return out;
}
const changedEntries = (before: Map<string, string>, after: Map<string, string>): string[] => {
  const names = new Set([...before.keys(), ...after.keys()]);
  return [...names].filter((n) => before.get(n) !== after.get(n)).sort();
};

// ---------------------------------------------------------------------------------------------- starting packages
const prefixRels = (xml: string): string => xml
  .replace(/<Relationships\b/, '<p:Relationships').replace(/<\/Relationships>/, '</p:Relationships>')
  .replace(/xmlns=/, 'xmlns:p=').replace(/<Relationship\b/g, '<p:Relationship');
const prefixTypes = (xml: string): string => xml
  .replace(/<Types\b/, '<t:Types').replace(/<\/Types>/, '</t:Types>').replace(/xmlns=/, 'xmlns:t=')
  .replace(/<(Default|Override)\b/g, '<t:$1');

async function twoCharts(): Promise<Buffer> {
  const s = await open(createBlankDocxBuffer('blank'));
  const base = {
    action: 'chart', operation: 'insert', anchor: s.ast.sections[0].nodes[0].id, chartType: 'column',
    categories: ['Q1', 'Q2', 'Q3'], series: [{ name: 'A', values: [1, 2, 3] }, { name: 'B', values: [3, 2, 1] }],
  };
  await s.run([{ ...base, id: 'k1', title: 'One' }, { ...base, id: 'k2', title: 'Two', chartType: 'line' }]);
  return snapshot(s);
}
function rewrite(buffer: Buffer, edit: (zip: AdmZip) => void): Buffer {
  const zip = new AdmZip(buffer);
  edit(zip);
  return zip.toBuffer();
}

const STARTS: Array<{ name: string; build: () => Promise<Buffer> }> = [
  { name: 'blank', build: async () => createBlankDocxBuffer('blank') },
  { name: 'word-made', build: async () => readFileSync(WORD_CHART) },
  { name: 'two charts', build: twoCharts },
  {
    name: 'prefixed document rels and content types',
    build: async () => rewrite(await twoCharts(), (zip) => {
      zip.updateFile('word/_rels/document.xml.rels', Buffer.from(prefixRels(zip.readAsText('word/_rels/document.xml.rels'))));
      zip.updateFile('[Content_Types].xml', Buffer.from(prefixTypes(zip.readAsText('[Content_Types].xml'))));
    }),
  },
  {
    name: 'prefixed chart rels',
    build: async () => rewrite(await twoCharts(), (zip) => {
      for (const e of zip.getEntries()) if (/^word\/charts\/_rels\/.*\.rels$/.test(e.entryName)) zip.updateFile(e.entryName, Buffer.from(prefixRels(zip.readAsText(e.entryName))));
    }),
  },
  {
    name: 'two charts sharing one workbook',
    build: async () => rewrite(await twoCharts(), (zip) => {
      const books = zip.getEntries().map((e) => e.entryName).filter((n) => /^word\/embeddings\/.*\.xlsx$/.test(n)).sort();
      const [first, second] = books as [string, string];
      const rels2 = 'word/charts/_rels/chart2.xml.rels';
      zip.updateFile(rels2, Buffer.from(zip.readAsText(rels2).replace(/Target="[^"]*embeddings\/[^"]*"/, `Target="../embeddings/${first.split('/').pop()}"`)));
      zip.deleteFile(second);
    }),
  },
];

// ---------------------------------------------------------------------------------------------- operation generator
const CHART_TYPES = ['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'columnStacked', 'barStacked', 'columnPercentStacked', 'combo', 'column3D', 'bar3D', 'line3D', 'pie3D'] as const;
const NAMES = ['North', 'South', 'East', 'West', 'Alpha', 'Beta'];

function dataset(r: Rand, scatter = false): { categories: string[]; series: Array<{ name: string; values: number[]; chartType?: string }> } {
  const n = int(r, 2, 5);
  const categories = scatter ? Array.from({ length: n }, (_, i) => String(i * 2 + int(r, 0, 1))) : Array.from({ length: n }, (_, i) => `C${i + 1}`);
  const count = int(r, 1, 3);
  const names = [...NAMES].sort(() => r() - 0.5).slice(0, count);
  return { categories, series: names.map((name) => ({ name, values: Array.from({ length: n }, () => int(r, -3, 40)) })) };
}

function randomStyle(r: Rand): Record<string, unknown> {
  const style: Record<string, unknown> = {};
  const keys = ['legend', 'dataLabels', 'axes', 'trendline', 'titleFormat', 'chartFont', 'bars', 'lines', 'chartArea', 'plotArea', 'series'];
  for (const key of keys) {
    if (!chance(r, 0.22)) continue;
    switch (key) {
      case 'legend': style.legend = { position: pick(r, ['b', 't', 'l', 'r', 'none', 'bad']), ...(chance(r, 0.3) ? { overlay: chance(r, 0.5) } : {}) }; break;
      case 'dataLabels': style.dataLabels = { show: chance(r, 0.6), ...(chance(r, 0.4) ? { position: pick(r, ['ctr', 'outEnd', 'inEnd', 'bestFit']) } : {}) }; break;
      case 'axes': style.axes = {
        ...(chance(r, 0.7) ? { category: { ...(chance(r, 0.5) ? { labelRotationDeg: int(r, -100, 95) } : {}), ...(chance(r, 0.5) ? { labelSizePt: int(r, 0, 20) } : {}), ...(chance(r, 0.4) ? { showAxisLine: chance(r, 0.5) } : {}), ...(chance(r, 0.3) ? { majorGridlines: chance(r, 0.5), gridlineColor: hex(r) } : {}) } } : {}),
        ...(chance(r, 0.7) ? { value: { ...(chance(r, 0.4) ? { gridlineColor: hex(r) } : {}), ...(chance(r, 0.4) ? { min: int(r, -5, 5) } : {}), ...(chance(r, 0.4) ? { max: int(r, 6, 50) } : {}), ...(chance(r, 0.3) ? { showAxisLine: chance(r, 0.5) } : {}), ...(chance(r, 0.3) ? { labelSizePt: int(r, 1, 18) } : {}) } } : {}),
      }; break;
      case 'trendline': style.trendline = { type: pick(r, ['linear', 'exponential', 'movingAvg', 'polynomial', 'none']), ...(chance(r, 0.4) ? { seriesNumbers: [int(r, 1, 4)] } : {}) }; break;
      case 'titleFormat': style.titleFormat = { ...(chance(r, 0.7) ? { sizePt: int(r, 8, 30) } : {}), ...(chance(r, 0.5) ? { bold: chance(r, 0.5) } : {}), ...(chance(r, 0.4) ? { color: hex(r) } : {}) }; break;
      case 'chartFont': style.chartFont = { ...(chance(r, 0.6) ? { family: pick(r, ['Arial', 'Georgia', 'Courier New']) } : {}), ...(chance(r, 0.6) ? { sizePt: int(r, 6, 16) } : {}), ...(chance(r, 0.4) ? { color: hex(r) } : {}) }; break;
      case 'bars': style.bars = { ...(chance(r, 0.7) ? { gapPercent: int(r, 0, 520) } : {}), ...(chance(r, 0.5) ? { overlapPercent: int(r, -100, 100) } : {}) }; break;
      case 'lines': style.lines = { ...(chance(r, 0.6) ? { widthPt: pick(r, [0.5, 2, 3.5, 25]) } : {}), ...(chance(r, 0.5) ? { smooth: chance(r, 0.5) } : {}), ...(chance(r, 0.5) ? { showMarkers: chance(r, 0.5) } : {}) }; break;
      case 'chartArea':
      case 'plotArea': style[key] = { ...(chance(r, 0.7) ? { fillColor: chance(r, 0.2) ? 'none' : hex(r) } : {}), ...(chance(r, 0.5) ? { borderColor: chance(r, 0.2) ? 'none' : hex(r) } : {}) }; break;
      case 'series': style.series = [{ seriesNumber: int(r, 1, 4), points: [{ pointNumber: int(r, 1, 6), color: hex(r) }] }]; break;
    }
  }
  return style;
}

interface Step { label: string; ops: unknown[]; targetsChart?: string; changesData: boolean; isUpdate: boolean }

function nextStep(r: Rand, s: Session, counter: { n: number }): Step {
  const charts = chartNodes(s);
  const choice = charts.length === 0 ? 'insert' : pick(r, ['insert', 'update', 'update', 'update', 'update', 'delete', 'paragraph']);
  if (choice === 'insert') {
    const type = pick(r, CHART_TYPES);
    const d = dataset(r, type === 'scatter');
    const series = type === 'combo' ? d.series.map((x, i) => ({ ...x, chartType: i === 0 ? 'column' : 'line' })) : d.series;
    const id = `f${++counter.n}`;
    const op: Record<string, unknown> = { action: 'chart', operation: 'insert', anchor: 'end', id, chartType: type, title: `Chart ${id}`, categories: d.categories, series };
    if (chance(r, 0.4)) op.style = randomStyle(r);
    if (chance(r, 0.3)) op.widthIn = pick(r, [3, 6, 7.5, 30]);
    if (chance(r, 0.2)) op.altText = `alt ${id}`;
    return { label: `insert ${type}`, ops: [op], changesData: false, isUpdate: false };
  }
  if (choice === 'paragraph') {
    return { label: 'insert paragraph', ops: [{ action: 'insert', type: 'paragraph', anchor: 'end', text: 'filler' }], changesData: false, isUpdate: false };
  }
  const target = pick(r, charts);
  if (choice === 'delete') {
    return { label: `delete ${target.id}`, ops: [{ action: 'delete', target: target.id }], changesData: false, isUpdate: false };
  }
  const op: Record<string, unknown> = { action: 'chart', operation: 'update', target: target.id };
  let changesData = false;
  const parts: string[] = [];
  const meta = target.metadata?.chartData ?? target.metadata ?? {};
  const currentSeries: Array<{ name: string; values: number[] }> = Array.isArray(meta.series) ? meta.series : [];
  const cats: string[] = Array.isArray(meta.categories) ? meta.categories : ['a', 'b', 'c'];
  const kind = pick(r, ['title', 'style', 'style', 'data', 'addSeries', 'removeSeries', 'type', 'size', 'removeTitle', 'mixed'] as const);
  if (kind === 'title' || kind === 'mixed') { op.title = `T${int(r, 1, 99)}`; parts.push('title'); }
  if (kind === 'style' || kind === 'mixed') { op.style = randomStyle(r); parts.push('style'); }
  if (kind === 'data') {
    const n = cats.length || 3;
    op.series = (currentSeries.length ? currentSeries : [{ name: 'A', values: [1, 2, 3] }]).map((x) => ({ name: x.name, values: Array.from({ length: n }, () => int(r, 0, 30)) }));
    changesData = true; parts.push('data');
  }
  if (kind === 'addSeries') {
    const n = cats.length || 3;
    op.series = [...currentSeries.map((x) => ({ name: x.name, values: x.values })), { name: `Extra${int(r, 1, 99)}`, values: Array.from({ length: n }, () => int(r, 0, 30)) }];
    changesData = true; parts.push('addSeries');
  }
  if (kind === 'removeSeries') {
    op.removeSeries = [currentSeries.length ? pick(r, currentSeries).name : 'none'];
    changesData = true; parts.push('removeSeries');
  }
  if (kind === 'type') { op.chartType = pick(r, CHART_TYPES); changesData = true; parts.push(`type ${String(op.chartType)}`); }
  if (kind === 'size') { op.widthIn = pick(r, [2, 5, 6.5, 40]); if (chance(r, 0.5)) op.heightIn = pick(r, [1.5, 2.6, 4]); parts.push('size'); }
  if (kind === 'removeTitle') { op.removeTitle = true; parts.push('removeTitle'); }
  if (kind !== 'title' && chance(r, 0.15)) op.altText = `alt${int(r, 1, 9)}`;
  if (chance(r, 0.35)) op.allowRebuild = true;
  return { label: `update ${target.id} [${parts.join(',')}]`, ops: [op], targetsChart: target.id, changesData, isUpdate: true };
}

// ---------------------------------------------------------------------------------------------- the test
describe('random chart operation sequences keep the package consistent', () => {
  for (const start of STARTS) {
    it(`${start.name}: ${SEEDS} seeds x ${STEPS} steps`, async () => {
      const buffer = await start.build();
      const baseline = inspectPackage(buffer);
      const baselineErrors = (await validateDocxBuffer(buffer)).errors?.length ?? 0;
      for (let seed = 1; seed <= SEEDS; seed++) {
        const r = rng(seed * 7919 + start.name.length);
        const s = await open(buffer);
        const counter = { n: 0 };
        const where = (step: number, label: string): string => `[start="${start.name}" seed=${seed} step=${step} ${label}]`;
        for (let step = 1; step <= STEPS; step++) {
          const plan = nextStep(r, s, counter);
          const before = snapshot(s);
          const beforeDigest = digest(before);
          const beforeCharts = chartNodes(s).length;
          let res: any;
          try {
            res = await s.run(plan.ops);
          } catch (error) {
            throw new Error(`${where(step, plan.label)} executeBatch threw: ${(error as Error).message}\nops: ${JSON.stringify(plan.ops)}`);
          }
          if (process.env.FUZZ_STATS) appendFileSync(process.env.FUZZ_STATS, `${res.success === false ? 'refused' : 'ok'}\t${plan.label.replace(/ f\d+| k\d| \w{8}\b/g, '').replace(/\[.*\]/, '[]')}\t${plan.isUpdate ? ((plan.ops[0] as any).style ? 'style' : 'other') : ''}\n`);
          const after = snapshot(s);
          const afterDigest = digest(after);
          const tag = `${where(step, plan.label)}\nops: ${JSON.stringify(plan.ops)}\nresult: ${JSON.stringify({ success: res.success, errors: res.errors }).slice(0, 400)}`;

          if (res.success === false) {
            expect(changedEntries(beforeDigest, afterDigest), `a refused batch changed the package\n${tag}`).toEqual([]);
            expect(chartNodes(s).length, `refused batch changed the chart count\n${tag}`).toBe(beforeCharts);
          } else if (plan.isUpdate) {
            const changed = changedEntries(beforeDigest, afterDigest);
            const chartXml = changed.filter((n) => /^word\/charts\/chart\d+\.xml$/.test(n));
            expect(chartXml.length, `an update changed more than one chart part: ${chartXml.join(', ')}\n${tag}`).toBeLessThanOrEqual(1);
            const books = changed.filter((n) => /^word\/embeddings\//.test(n));
            const rebuildAllowed = (plan.ops[0] as { allowRebuild?: boolean }).allowRebuild === true;
            if (!plan.changesData) expect(books, `a no-data update touched a workbook\n${tag}`).toEqual([]);
            else if (!rebuildAllowed) expect(books.length, `a data update changed several workbooks\n${tag}`).toBeLessThanOrEqual(1);
          }

          const report = inspectPackage(after);
          const fresh = report.problems.filter((p) => !baseline.problems.includes(p));
          expect(fresh, `package invariant broken\n${tag}`).toEqual([]);
          const nodes = chartNodes(s);
          expect(report.chartParts.length, `chart parts vs chart nodes\n${tag}`).toBe(nodes.length);
          expect(report.chartRelCount, `chart relationships vs chart nodes\n${tag}`).toBe(nodes.length);

          if (step % VALIDATE_EVERY === 0 || step === STEPS) {
            const verdict = await validateDocxBuffer(after);
            expect(verdict.errors?.length ?? 0, `validator errors grew\n${JSON.stringify(verdict.errors?.slice(0, 3))}\n${tag}`).toBeLessThanOrEqual(baselineErrors);
            const reopened = await open(after);
            expect(chartNodes(reopened).length, `reopened package lost or gained charts\n${tag}`).toBe(nodes.length);
          }
        }
      }
    }, 600_000);
  }
});
