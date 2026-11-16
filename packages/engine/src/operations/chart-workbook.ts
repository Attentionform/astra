/**
 * Workbook policy for chart updates (0.5.3, DN1 section "workbook policy").
 *
 * The embedded workbook behind a chart is the user's data. ASTra writes it only
 * when the update changes chart data, and only when it can prove it replaces its
 * own generated workbook (or the caller opts in with `allowRebuild`).
 *
 *   Rule 1: an update that carries no data change never touches the workbook.
 *   Rule 2: a data change rewrites the workbook only if the stored workbook is
 *           part-by-part identical to what ASTra would generate from the chart's
 *           stored data. Anything else (Excel edits, extra sheets, formulas,
 *           a workbook shared with another chart) is refused unless the caller
 *           passes `allowRebuild: true`. A shared workbook is then copied to a
 *           private part first; the other chart keeps the original.
 *   A linked (external) workbook is never rewritten.
 */
import AdmZip from '../adm-zip-safe';
import { generateEmbeddedWorkbook } from '../manipulators/chart3d';
import type { DocxZipAdapter } from '../parser/types';

export interface StoredChartData {
  categories: string[];
  series: Array<{ name: string; values: number[]; xValues?: number[] }>;
  /** True when the chart stores one x list per series (scatter layout). */
  scatter: boolean;
}

export interface WorkbookRefusal {
  message: string;
  remedy: string;
}

export type WorkbookDecision =
  | { action: 'none'; reason: 'no-data-change' | 'no-external-data' }
  | {
      action: 'rewrite';
      /** Part that receives the new workbook bytes. */
      path: string;
      /** True when the stored workbook was recognised as ASTra-generated. */
      recognised: boolean;
      /** Set when a shared workbook must first be copied to a private part. */
      privateCopy?: { relsPath: string; relId: string; newPath: string; previousPath: string };
    }
  | { action: 'refuse'; refusal: WorkbookRefusal };

interface RelEntry {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  return m ? (m[1] ?? m[2]) : undefined;
}

function parseRels(xml: string): RelEntry[] {
  const out: RelEntry[] = [];
  const re = /<Relationship\b[^>]*>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const id = attr(m[0], 'Id');
    const target = attr(m[0], 'Target');
    if (id === undefined || target === undefined) continue;
    out.push({
      id,
      type: attr(m[0], 'Type') ?? '',
      target,
      external: (attr(m[0], 'TargetMode') ?? '').toLowerCase() === 'external',
    });
  }
  return out;
}

/** Resolves a relationship target against the directory of its source part. */
function resolveTarget(sourceDir: string, target: string): string {
  const parts = (target.startsWith('/') ? target.slice(1) : `${sourceDir}${target}`).split('/');
  const out: string[] = [];
  for (const part of parts) {
    if (part === '' || part === '.') continue;
    if (part === '..') out.pop();
    else out.push(part);
  }
  return out.join('/');
}

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash >= 0 ? partPath.slice(0, slash + 1) : '';
  const base = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return `${dir}_rels/${base}.rels`;
}

function sourceDirOf(relsPath: string): string {
  // a/_rels/b.xml.rels -> a/
  const idx = relsPath.lastIndexOf('_rels/');
  return idx > 0 ? relsPath.slice(0, idx) : '';
}

function entriesOf(buffer: Buffer): Map<string, Buffer> | null {
  try {
    const zip = new AdmZip(buffer);
    const map = new Map<string, Buffer>();
    for (const entry of zip.getEntries()) {
      if (entry.isDirectory) continue;
      map.set(entry.entryName, entry.getData());
    }
    return map;
  } catch {
    return null;
  }
}

function sameParts(a: Map<string, Buffer>, b: Map<string, Buffer>): boolean {
  if (a.size !== b.size) return false;
  for (const [name, data] of a) {
    const other = b.get(name);
    if (!other || !data.equals(other)) return false;
  }
  return true;
}

/** True when `actual` equals a workbook ASTra would generate from the stored chart data. */
export function isGeneratedWorkbook(actual: Buffer, stored: StoredChartData): boolean {
  const have = entriesOf(actual);
  if (!have) return false;
  const rowCount = Math.max(
    stored.categories.length,
    ...stored.series.map((s) => s.values.length),
    ...stored.series.map((s) => s.xValues?.length ?? 0),
    0
  );
  const padded = Array.from({ length: rowCount }, (_, i) => stored.categories[i] ?? '');
  const series = stored.series.map((s) => ({ name: s.name, values: s.values }));
  const scatterX = stored.scatter
    ? stored.series.map((s) => s.xValues ?? Array.from({ length: s.values.length }, (_, i) => {
        const numeric = Number(stored.categories[i]);
        return Number.isFinite(numeric) ? numeric : i;
      }))
    : undefined;
  const variants: string[][] = [stored.categories];
  if (padded.length !== stored.categories.length) variants.push(padded);
  for (const categories of variants) {
    let rebuilt: Buffer;
    try {
      rebuilt = generateEmbeddedWorkbook(series, categories, scatterX);
    } catch {
      continue;
    }
    const want = entriesOf(rebuilt);
    if (want && sameParts(have, want)) return true;
  }
  return false;
}

function uniqueWorkbookPath(names: ReadonlySet<string>): string {
  for (let n = 1; ; n++) {
    const candidate = `word/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
    if (!names.has(candidate)) return candidate;
  }
}

export interface WorkbookPlanInput {
  zip: DocxZipAdapter;
  chartPath: string;
  externalDataRId?: string;
  stored: StoredChartData;
  dataChange: boolean;
  allowRebuild: boolean;
}

export function planWorkbook(input: WorkbookPlanInput): WorkbookDecision {
  const { zip, chartPath, stored } = input;
  if (!input.dataChange) return { action: 'none', reason: 'no-data-change' };
  if (!input.externalDataRId) return { action: 'none', reason: 'no-external-data' };
  if (!zip.readBuffer || !zip.entryNames) {
    return {
      action: 'refuse',
      refusal: {
        message: `E_VALIDATION: this package adapter cannot inspect the workbook behind ${chartPath}.`,
        remedy: 'Remedy: update the chart in a document opened by the engine parser.',
      },
    };
  }
  const relsPath = relsPathFor(chartPath);
  const relsXml = zip.readTextOrNull(relsPath);
  const rel = relsXml ? parseRels(relsXml).find((r) => r.id === input.externalDataRId) : undefined;
  if (!rel) {
    return {
      action: 'refuse',
      refusal: {
        message: `E_VALIDATION: ${chartPath} points at workbook relationship ${input.externalDataRId}, which is missing from ${relsPath}.`,
        remedy: 'Remedy: delete the chart and insert it again; ASTra does not repair a broken workbook relationship.',
      },
    };
  }
  if (rel.external) {
    return {
      action: 'refuse',
      refusal: {
        message: `E_VALIDATION: ${chartPath} reads its data from a linked workbook (${rel.target}). ASTra never rewrites a linked workbook, and the chart cache would disagree with it.`,
        remedy: 'Remedy: change the data in the linked workbook, or delete the chart and insert a new one with the data.',
      },
    };
  }
  const workbookPath = resolveTarget(sourceDirOf(relsPath), rel.target);

  // Other relationships that point at the same workbook part make it shared.
  const names = new Set(zip.entryNames());
  const sharers: string[] = [];
  for (const name of names) {
    if (!/\.rels$/i.test(name)) continue;
    const xml = zip.readTextOrNull(name);
    if (!xml) continue;
    const dir = sourceDirOf(name);
    for (const other of parseRels(xml)) {
      if (other.external) continue;
      if (name === relsPath && other.id === rel.id) continue;
      if (resolveTarget(dir, other.target) === workbookPath) {
        sharers.push(name.replace(/_rels\//, '').replace(/\.rels$/i, ''));
      }
    }
  }

  const bytes = names.has(workbookPath) ? zip.readBuffer(workbookPath) : null;
  const recognised = bytes !== null && isGeneratedWorkbook(bytes, stored);

  if (sharers.length > 0) {
    if (!input.allowRebuild) {
      return {
        action: 'refuse',
        refusal: {
          message: `E_VALIDATION: the workbook ${workbookPath} behind ${chartPath} is also used by ${sharers.join(', ')}. A data change would rewrite data the other chart reads.`,
          remedy: 'Remedy: pass allowRebuild:true to give this chart its own private copy of the workbook (the other chart keeps the original), or leave the data unchanged.',
        },
      };
    }
    const newPath = uniqueWorkbookPath(names);
    return {
      action: 'rewrite',
      path: newPath,
      recognised,
      privateCopy: { relsPath, relId: rel.id, newPath, previousPath: workbookPath },
    };
  }

  if (!recognised && !input.allowRebuild) {
    const why = bytes === null
      ? `the workbook part ${workbookPath} is missing`
      : `the workbook ${workbookPath} is not the workbook ASTra generated for this chart (it was edited, or built by another application)`;
    return {
      action: 'refuse',
      refusal: {
        message: `E_VALIDATION: ${why}. A data change would replace it with a generated workbook and lose its other content (extra sheets, formulas, formatting, named ranges).`,
        remedy: 'Remedy: pass allowRebuild:true to replace the workbook with a generated one, or change only formatting (title, style, size) which never touches the workbook.',
      },
    };
  }
  return { action: 'rewrite', path: workbookPath, recognised };
}

/** Writes the new workbook bytes according to the decision. */
export function applyWorkbookDecision(zip: DocxZipAdapter, decision: WorkbookDecision, bytes: Buffer): void {
  if (decision.action !== 'rewrite') return;
  if (decision.privateCopy) {
    const copy = decision.privateCopy;
    const relsXml = zip.readText(copy.relsPath);
    const next = relsXml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      if (attr(tag, 'Id') !== copy.relId) return tag;
      const target = attr(tag, 'Target');
      if (target === undefined) return tag;
      const base = copy.newPath.split('/').pop() as string;
      return tag.replace(/(\sTarget\s*=\s*)(?:"[^"]*"|'[^']*')/, `$1"../embeddings/${base}"`);
    });
    zip.writeText(copy.relsPath, next);
  }
  zip.writeBuffer(decision.path, bytes);
}
