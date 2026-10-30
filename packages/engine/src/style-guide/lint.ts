/**
 * Advisory style linter: read-only AST pass that reports taste defects
 * without ever failing the save.
 *
 * "Style is advisory, schema is law" — every finding lands in the returned
 * array; this module never throws for document content (best-effort per
 * node, `E_VALIDATION` stays reserved for schema/structural failures in the
 * save path itself). Runs only when a style guide is active; without one
 * there is no house style to advise against, so the result is empty.
 */
import type { DocumentAST, Node } from '../ast';
import type { StyleGuideConfig } from './types';
import {
  contrastRatio,
  TABLE_CONTRAST_MIN_RATIO,
} from '../operations/table';

export interface StyleWarning {
  nodeId: string;
  rule: string;
  message: string;
}

export const STYLE_WARNING_RULES = [
  'table-header-contrast',
  'chart-duplicate-colors',
  'table-orphan',
] as const;

export type StyleWarningRule = (typeof STYLE_WARNING_RULES)[number];

function directChildren(parent: Element, tagName: string): Element[] {
  const out: Element[] = [];
  const kids = parent.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const child = kids[i] as Element;
    if (
      (child as unknown as { nodeType?: number }).nodeType === 1 &&
      (child as Element).tagName === tagName
    ) {
      out.push(child as Element);
    }
  }
  return out;
}

function cellFillOf(tc: Element): string | null {
  const tcPr = directChildren(tc, 'w:tcPr')[0];
  const shd = tcPr ? directChildren(tcPr, 'w:shd')[0] : undefined;
  const v = shd?.getAttribute('w:fill');
  return v && /^[0-9A-Fa-f]{6}$/.test(v) ? v.toUpperCase() : null;
}

/**
 * First explicit 6-digit text color in a cell's OWN paragraphs (direct w:p
 * children only — never descends into a nested table), or null when the
 * cell relies on the default. Mirrors the table op guardrail: absent color
 * renders as black on a light page, so callers treat null as '000000'.
 */
function firstExplicitCellTextColor(tc: Element): string | null {
  for (const p of directChildren(tc, 'w:p')) {
    const runs: Element[] = [...directChildren(p, 'w:r')];
    for (const tag of ['w:hyperlink', 'w:ins', 'w:del']) {
      for (const wrap of directChildren(p, tag)) {
        const wRuns = wrap.getElementsByTagName('w:r');
        for (let i = 0; i < wRuns.length; i++) runs.push(wRuns[i] as Element);
      }
    }
    for (const r of runs) {
      const rPr = directChildren(r, 'w:rPr')[0];
      const colorEl = rPr?.getElementsByTagName('w:color')[0] as
        | Element
        | undefined;
      const v = colorEl?.getAttribute('w:val');
      if (v && /^[0-9A-Fa-f]{6}$/.test(v)) return v.toUpperCase();
    }
  }
  return null;
}

function lintTableContrast(node: Node, guide: StyleGuideConfig): StyleWarning | null {
  try {
    const xml = node.xml as Element | undefined;
    if (!xml || typeof (xml as Element).getElementsByTagName !== 'function') {
      return null;
    }
    const rows = directChildren(xml, 'w:tr');
    if (rows.length === 0) return null;
    const recommended = (guide.tables.headerText ?? 'FFFFFF')
      .replace(/^#/, '')
      .toUpperCase();
    for (const tc of directChildren(rows[0], 'w:tc')) {
      const fill = cellFillOf(tc);
      if (!fill) continue;
      const text = firstExplicitCellTextColor(tc) ?? '000000';
      if (contrastRatio(text, fill) < TABLE_CONTRAST_MIN_RATIO) {
        return {
          nodeId: node.id,
          rule: 'table-header-contrast',
          message:
            `Header cell has fill '${fill}' but text color is '${text}' ` +
            `(contrast ${contrastRatio(text, fill).toFixed(2)}:1 below WCAG AA ` +
            `${TABLE_CONTRAST_MIN_RATIO}:1). Recommended: '${recommended}'.`,
        };
      }
    }
    return null;
  } catch {
    return null;
  }
}

function lintTableOrphan(node: Node, guide?: StyleGuideConfig): StyleWarning | null {
  try {
    if (guide?.tables?.keepTogether === false) return null;
    const xml = node.xml as Element | undefined;
    if (!xml || typeof (xml as Element).getElementsByTagName !== 'function') {
      return null;
    }
    const rows = directChildren(xml, 'w:tr');
    if (rows.length === 0) return null;
    let missing = 0;
    for (const tr of rows) {
      const trPr = directChildren(tr, 'w:trPr')[0];
      if (!trPr || trPr.getElementsByTagName('w:cantSplit').length === 0) {
        missing++;
      }
    }
    if (missing === 0) return null;
    return {
      nodeId: node.id,
      rule: 'table-orphan',
      message:
        `Table has ${missing} row(s) without cantSplit and may split across ` +
        `page boundaries. Remedy: run { action:'table', target:'${node.id}', ` +
        `operation:'setKeepTogether' } to keep rows together.`,
    };
  } catch {
    return null;
  }
}

function seriesColorsFromChartXml(chartXml: string): string[] {
  if (typeof chartXml !== 'string' || chartXml.length === 0) return [];
  const out: string[] = [];
  try {
    const blocks = Array.from(chartXml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g));
    for (const m of blocks) {
      const body = (m as RegExpMatchArray)[1] ?? '';
      const fill = /<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/.exec(body);
      out.push(fill ? (fill[1] as string).toUpperCase() : '');
    }
  } catch {
    return [];
  }
  return out;
}

/**
 * Builds a persisted chart-part reader for a loaded document. The returned
 * function maps a chart node id to its stored chart XML (or undefined when
 * the part is missing), letting the linter compare the colors renderers use
 * rather than the live-only insert detail.
 */
export function chartXmlGetterForAst(
  ast: DocumentAST,
  readText: (partName: string) => string | null | undefined,
): (nodeId: string) => string | undefined {
  const byId = new Map<string, Node>();
  try {
    for (const section of (ast as DocumentAST | undefined)?.sections ?? []) {
      for (const node of section?.nodes ?? []) {
        if (node && typeof node.id === 'string') byId.set(node.id, node as Node);
      }
    }
  } catch {
    /* empty map keeps the getter total */
  }
  return (nodeId: string): string | undefined => {
    try {
      const node = byId.get(nodeId);
      const xml = node?.xml as unknown as Element | undefined;
      if (!xml || typeof (xml as Element).getElementsByTagName !== 'function') {
        return undefined;
      }
      const chartEls = (xml as Element).getElementsByTagName('c:chart');
      if (chartEls.length === 0) return undefined;
      const rId = (chartEls[0] as Element).getAttribute('r:id');
      if (!rId) return undefined;
      const rel = (ast as DocumentAST).relationships?.[rId];
      const target = rel?.target;
      if (!target) return undefined;
      const part = target.startsWith('word/') ? target : `word/${target}`;
      const text = readText(part);
      return typeof text === 'string' && text.length > 0 ? text : undefined;
    } catch {
      return undefined;
    }
  };
}

function seriesColorsOf(
  node: Node,
  getChartXml?: (nodeId: string) => string | undefined,
): string[] {
  if (getChartXml) {
    try {
      const stored = getChartXml(node.id);
      if (typeof stored === 'string' && stored.length > 0) {
        const fromStored = seriesColorsFromChartXml(stored);
        if (fromStored.some((c) => c !== '')) return fromStored;
      }
    } catch {
      /* fall through to the live detail below */
    }
  }
  const meta = node.metadata as unknown as Record<string, unknown>;
  const raw =
    (meta['series'] as Array<{ color?: unknown }> | undefined) ??
    (meta['chartData'] as Record<string, unknown> | undefined)?.['series'] as
      | Array<{ color?: unknown }>
      | undefined;
  if (!Array.isArray(raw)) return [];
  return raw.map((s) =>
    typeof s?.color === 'string' ? s.color.replace(/^#/, '').toUpperCase() : '',
  );
}

function lintChartDuplicates(
  node: Node,
  getChartXml?: (nodeId: string) => string | undefined,
): StyleWarning | null {
  try {
    const colors = seriesColorsOf(node, getChartXml);
    const seen = new Map<string, number>();
    for (let i = 0; i < colors.length; i++) {
      const c = colors[i];
      if (!c) continue;
      if (seen.has(c)) {
        return {
          nodeId: node.id,
          rule: 'chart-duplicate-colors',
          message:
            `Chart series ${seen.get(c)} and ${i} share color '${c}'. ` +
            `Remedy: give each series a distinct color from the guide chart palette.`,
        };
      }
      seen.set(c, i);
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Advisory style pass over the live AST. Returns warnings only — never
 * throws for document content (a corrupt node yields no warning, not a
 * failed save). Empty array when no guide is active.
 */
export function lintDocumentStyle(
  ast: DocumentAST,
  guide?: StyleGuideConfig,
  getChartXml?: (nodeId: string) => string | undefined,
): StyleWarning[] {
  const out: StyleWarning[] = [];
  try {
    if (!guide) return out;
    const sections = (ast as DocumentAST | undefined)?.sections;
    if (!Array.isArray(sections)) return out;
    for (const section of sections) {
      for (const node of section?.nodes ?? []) {
        if (!node || typeof node.id !== 'string') continue;
        if (node.type === 'table') {
          const contrast = lintTableContrast(node, guide);
          if (contrast) out.push(contrast);
          const orphan = lintTableOrphan(node, guide);
          if (orphan) out.push(orphan);
        } else if (node.type === 'chart' || node.type === 'chart3d') {
          const dup = lintChartDuplicates(node, getChartXml);
          if (dup) out.push(dup);
        }
      }
    }
  } catch {
    /* advisory only — never fail the caller */
  }
  return out;
}
