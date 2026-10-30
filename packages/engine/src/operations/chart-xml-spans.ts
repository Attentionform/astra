import { findChartNamespaceElementsScoped } from '../parser/factory';

export interface Span {
  start: number;
  end: number;
}

export type ChartNamespaceScope = ReadonlyMap<string, string>;

export function escapeXml(value: string): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Index of the '>' that closes the tag starting at `start` (quote-aware). */
export function tagEnd(xml: string, start: number): number {
  let quote: string | null = null;
  for (let i = start; i < xml.length; i++) {
    const c = xml[i];
    if (quote) {
      if (c === quote) quote = null;
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === '>') {
      return i;
    }
  }
  return -1;
}

export function tagNameOf(startTagXml: string): string {
  const m = /^<([^\s/>]+)/.exec(startTagXml);
  return m ? m[1] : '';
}

export function elementSpanAt(xml: string, tag: string, start: number): Span | null {
  const gt = tagEnd(xml, start);
  if (gt < 0) return null;
  if (xml[gt - 1] === '/') return { start, end: gt + 1 };
  const re = new RegExp(`<${tag}(?=[\\s/>])|</${tag}>`, 'g');
  re.lastIndex = start;
  let depth = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    if (m[0][1] === '/') {
      depth--;
      if (depth === 0) return { start, end: m.index + m[0].length };
      if (depth < 0) return null;
    } else {
      depth++;
    }
    if (depth > 512) return null;
  }
  return null;
}

export function findElement(xml: string, tag: string, from = 0, limit = xml.length): Span | null {
  const re = new RegExp(`<${tag}(?=[\\s/>])`, 'g');
  re.lastIndex = from;
  const m = re.exec(xml);
  if (!m || m.index >= limit) return null;
  return elementSpanAt(xml, tag, m.index);
}

export function findAllElements(xml: string, tag: string, from = 0, limit = xml.length): Span[] {
  const out: Span[] = [];
  let cursor = from;
  while (cursor < limit) {
    const span = findElement(xml, tag, cursor, limit);
    if (!span) break;
    out.push(span);
    cursor = span.end;
  }
  return out;
}

export function chartSpans(
  xml: string,
  localName: string,
  scope: ChartNamespaceScope,
  from = 0,
  limit = xml.length
): Span[] {
  return findChartNamespaceElementsScoped(xml, scope, [localName])
    .filter((element) => element.end >= 0 && element.start >= from && element.start < limit)
    .map((element) => ({ start: element.start, end: element.end }));
}

export function findChartElement(
  xml: string,
  localName: string,
  scope: ChartNamespaceScope,
  from = 0,
  limit = xml.length
): Span | null {
  return chartSpans(xml, localName, scope, from, limit)[0] ?? null;
}

export function findChartDirectChild(
  parentXml: string,
  localName: string,
  scope: ChartNamespaceScope
): Span | null {
  const element = findChartNamespaceElementsScoped(parentXml, scope, [localName])
    .find((candidate) => candidate.depth === 1 && candidate.end >= 0);
  return element ? { start: element.start, end: element.end } : null;
}

/** Every direct child with the given chart local name, in document order. */
export function findChartDirectChildren(
  parentXml: string,
  localName: string,
  scope: ChartNamespaceScope
): Span[] {
  return findChartNamespaceElementsScoped(parentXml, scope, [localName])
    .filter((candidate) => candidate.depth === 1 && candidate.end >= 0)
    .map((element) => ({ start: element.start, end: element.end }));
}

/**
 * Chart-level title: a direct chart-namespace <c:title> child of <c:chart>
 * before <c:plotArea>. A nested <c:title> (for example in a hostile extLst)
 * is not a legal CT_Chart title and must be ignored, matching the reader.
 */
export function findDirectChartTitleSpan(
  xml: string,
  chartSpan: Span,
  scope: ChartNamespaceScope
): Span | null {
  const block = xml.slice(chartSpan.start, chartSpan.end);
  const title = findChartDirectChild(block, 'title', scope);
  if (!title) return null;
  const plotArea = findChartDirectChild(block, 'plotArea', scope);
  if (plotArea && plotArea.start < title.start) return null;
  return { start: chartSpan.start + title.start, end: chartSpan.start + title.end };
}

export function chartPrefix(parentXml: string): string {
  const parent = tagNameOf(parentXml);
  const colon = parent.indexOf(':');
  return colon >= 0 ? parent.slice(0, colon + 1) : '';
}

/** Inserts a sibling/child using the parent's actual chart prefix or default namespace. */
export function chartQualifiedName(parentXml: string, localName: string): string {
  return chartPrefix(parentXml) + localName;
}

/**
 * First direct child with the given tag. Walks top-level children only, so a
 * nested node with the same tag (per-point `spPr`, per-series `shape`) can
 * never be mistaken for the series- or chart-level one.
 */
export function findDirectChild(parentXml: string, childTag: string): Span | null {
  const openEnd = tagEnd(parentXml, 0) + 1;
  const closeStart = parentXml.lastIndexOf('</');
  let cursor = openEnd;
  while (cursor < closeStart) {
    const re = /<([A-Za-z0-9:_-]+)(?=[\s/>])/g;
    re.lastIndex = cursor;
    const m = re.exec(parentXml);
    if (!m || m.index >= closeStart) return null;
    const span = elementSpanAt(parentXml, m[1], m.index);
    if (!span) return null;
    if (m[1] === childTag) return span;
    cursor = span.end;
  }
  return null;
}

export function attrValue(elementXml: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(elementXml);
  if (!m) return undefined;
  return m[1] !== undefined ? m[1] : m[2];
}

export function setAttribute(elementXml: string, name: string, value: string): string {
  const gt = tagEnd(elementXml, 0);
  if (gt < 0) return elementXml;
  const open = elementXml.slice(0, gt);
  const tail = elementXml.slice(gt);
  const re = new RegExp(`(\\s${name}\\s*=\\s*)(?:"[^"]*"|'[^']*')`);
  if (re.test(open)) return open.replace(re, `$1"${value}"`) + tail;
  const selfClose = /\/\s*$/.exec(open);
  if (selfClose) {
    const at = selfClose.index;
    return `${open.slice(0, at)} ${name}="${value}"${open.slice(at)}${tail}`;
  }
  return `${open} ${name}="${value}"${tail}`;
}

/** Replaces an element's text content, keeping its start-tag attributes. */
export function setElementText(elementXml: string, text: string): string {
  const gt = tagEnd(elementXml, 0);
  if (gt < 0) return elementXml;
  const selfClosed = elementXml[gt - 1] === '/';
  const open = selfClosed ? elementXml.slice(0, gt - 1) : elementXml.slice(0, gt);
  const keepSpace = /^\s|\s$/.test(text) && !/xml:space\s*=/.test(open);
  const close = `</${tagNameOf(elementXml)}>`;
  return `${open}${keepSpace ? ' xml:space="preserve"' : ''}>${escapeXml(text)}${close}`;
}
