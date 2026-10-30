/**
 * P2 loop-back: sniff a loaded chart part (<c:chartSpace> XML) instead of
 * hardcoding a 3D label. The actual plot element decides the family: explicit
 * 2D tags (c:barChart/c:lineChart/c:pieChart/c:doughnutChart/c:areaChart/
 * c:scatterChart) win
 * over a stray <c:view3D>, then explicit 3D tags (c:bar3DChart/
 * c:line3DChart/c:pie3DChart), then a generic 3DChart/view3D fallback.
 * 2D c:barChart splits on barDir + grouping: col/bar × clustered|stacked|
 * percentStacked → column|bar|columnStacked|barStacked|columnPercentStacked|
 * barPercentStacked. Truly undetectable XML falls back to 2D 'column' —
 * never a fabricated bar3D.
 */
import { BAR_2D_VARIANTS } from '../manipulators/chart2d-types';
import { CHART_NS } from '../ooxml/ns';

export const PLOT_LOCAL_NAMES_2D = ['barChart', 'lineChart', 'pieChart', 'doughnutChart', 'areaChart', 'scatterChart'] as const;
export const PLOT_LOCAL_NAMES_3D = ['bar3DChart', 'line3DChart', 'pie3DChart'] as const;
export const PLOT_LOCAL_NAMES_UNSUPPORTED = ['area3DChart', 'surfaceChart', 'surface3DChart', 'bubbleChart', 'stockChart', 'ofPieChart', 'radarChart'] as const;
export const PLOT_LOCAL_NAMES_ALL = [...PLOT_LOCAL_NAMES_2D, ...PLOT_LOCAL_NAMES_3D, ...PLOT_LOCAL_NAMES_UNSUPPORTED] as const;

/** Any XML namespace prefix, or none. */
const ANY_PREFIX = '(?:[A-Za-z_][\\w.-]*:)?';

export interface ChartNamespaceElement {
  localName: string;
  qualifiedName: string;
  /** Index of the element's '<' in the source string. */
  start: number;
  /** Index just past the matching closing tag (or the self-closing tag). */
  end: number;
  selfClosing: boolean;
  /** Nesting depth: 0 for a scanned root element, 1 for its direct children. */
  depth: number;
  /** In-scope namespace bindings at the element's start tag. */
  scope: ReadonlyMap<string, string>;
}

/** Index of the '>' closing the tag at `start` (quote-aware). */
function tagEnd(xml: string, start: number): number {
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

/** XML entities that may appear in an attribute value without a DTD. */
const NAMED_ENTITIES: Record<string, string> = {
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  amp: '&',
};

/**
 * Resolves the entity references an XML parser would apply to an attribute
 * value, so a namespace URI spelled with numeric character references
 * (`&#x74;`, `&#116;`) still compares equal to the literal chart namespace.
 * One pass keeps `&amp;#x74;` literal, as XML does.
 */
function decodeNamespaceUri(value: string): string {
  return value.replace(/&(?:#([0-9]+)|#x([0-9A-Fa-f]+)|(lt|gt|quot|apos|amp));/g, (match, dec, hex, name) => {
    if (name) return NAMED_ENTITIES[name];
    const code = dec !== undefined ? parseInt(dec, 10) : parseInt(hex, 16);
    if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return match;
    return String.fromCodePoint(code);
  });
}

/**
 * Elements whose namespace URI is the DrawingML chart namespace. Prefixes are
 * resolved against the in-scope xmlns declarations (including rebinding and
 * default namespaces), so vendor elements inside c:extLst are never mistaken
 * for plots even when they reuse a chart local name. Pass local names to
 * filter; omit to collect every chart-namespace element.
 */
export function findChartNamespaceElements(
  xml: string,
  localNames?: readonly string[]
): ChartNamespaceElement[] {
  const wanted = localNames ? new Set(localNames) : undefined;
  const found: ChartNamespaceElement[] = [];
  const scopes: Array<Map<string, string>> = [new Map()];
  const open: Array<ChartNamespaceElement | undefined> = [];
  let cursor = 0;
  while (cursor < xml.length) {
    const lt = xml.indexOf('<', cursor);
    if (lt < 0) break;
    if (xml.startsWith('<!--', lt)) {
      const end = xml.indexOf('-->', lt + 4);
      cursor = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt + 9);
      cursor = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<?', lt)) {
      const end = xml.indexOf('?>', lt + 2);
      cursor = end < 0 ? xml.length : end + 2;
      continue;
    }
    if (xml.startsWith('</', lt)) {
      const end = xml.indexOf('>', lt + 2);
      if (end < 0) break;
      const ref = open.pop();
      if (ref) ref.end = end + 1;
      if (scopes.length > 1) scopes.pop();
      cursor = end + 1;
      continue;
    }
    if (xml.startsWith('<!', lt)) {
      const end = tagEnd(xml, lt);
      if (end < 0) break;
      cursor = end + 1;
      continue;
    }
    const end = tagEnd(xml, lt);
    if (end < 0) break;
    const raw = xml.slice(lt + 1, end);
    const nameMatch = /^([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)/.exec(raw);
    if (!nameMatch) {
      cursor = end + 1;
      continue;
    }
    const qualifiedName = nameMatch[1];
    const selfClosing = /\/\s*$/.test(raw);
    const scope = new Map(scopes[scopes.length - 1]);
    const declRe = /\sxmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let decl: RegExpExecArray | null;
    while ((decl = declRe.exec(raw)) !== null) {
      scope.set(decl[1] ?? '', decodeNamespaceUri(decl[2] ?? decl[3] ?? ''));
    }
    const colon = qualifiedName.indexOf(':');
    const prefix = colon >= 0 ? qualifiedName.slice(0, colon) : '';
    const localName = colon >= 0 ? qualifiedName.slice(colon + 1) : qualifiedName;
    let ref: ChartNamespaceElement | undefined;
    if (scope.get(prefix) === CHART_NS && (!wanted || wanted.has(localName))) {
      ref = {
        localName,
        qualifiedName,
        start: lt,
        end: selfClosing ? end + 1 : -1,
        selfClosing,
        depth: open.length,
        scope,
      };
      found.push(ref);
    }
    if (!selfClosing) {
      open.push(ref);
      scopes.push(scope);
    }
    cursor = end + 1;
  }
  return found;
}

/** Chart-namespace prefix bindings declared anywhere in `xml`. */
export function chartNamespaceScopeFromXml(xml: string): Map<string, string> {
  const scope = new Map<string, string>();
  const re = /\sxmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const uri = decodeNamespaceUri(m[2] ?? m[3] ?? '');
    if (uri === CHART_NS) scope.set(m[1] ?? '', CHART_NS);
  }
  return scope;
}

function escapeNamespaceAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Scans a fragment with the inherited namespace bindings supplied by the full
 * part, so a substring (series, title, cache) resolves prefixes declared on an
 * ancestor outside the fragment. Positions and depths are fragment-relative.
 */
export function findChartNamespaceElementsScoped(
  xml: string,
  scope: ReadonlyMap<string, string>,
  localNames?: readonly string[]
): ChartNamespaceElement[] {
  const declarations: string[] = [];
  for (const [prefix, uri] of scope) {
    declarations.push(`${prefix ? `xmlns:${prefix}` : 'xmlns'}="${escapeNamespaceAttribute(uri)}"`);
  }
  const open = declarations.length > 0 ? `<astra-scope ${declarations.join(' ')}>` : '<astra-scope>';
  const wrapped = `${open}${xml}</astra-scope>`;
  const offset = open.length;
  return findChartNamespaceElements(wrapped, localNames).map((element) => ({
    ...element,
    start: element.start - offset,
    end: element.end - offset,
    depth: element.depth - 1,
  }));
}

/** Chart-namespace descendants of a DOM element, in document order. */
export function chartElementDescendants(parent: Element, localName: string): Element[] {
  const out: Element[] = [];
  try {
    const elements = parent.getElementsByTagName('*');
    for (let i = 0; i < elements.length; i++) {
      const el = elements[i] as Element;
      const name = el.localName || (el.nodeName ? el.nodeName.split(':').pop() : '');
      if (name === localName && el.namespaceURI === CHART_NS) out.push(el);
    }
  } catch {
    /* best-effort */
  }
  return out;
}

/**
 * Chart-level <c:title> only: a direct chart-namespace <c:title> child of
 * <c:chart> before <c:plotArea> — the same bound the writer uses when it
 * removes a title. Axis titles also use <c:title> but live inside the plot
 * area, so an unbounded descendant search reads an axis title as the chart
 * title once the chart-level one is cleared. Namespace URI + local name, so
 * default-namespace, prefixed and numerically prefixed parts all resolve.
 */
export function chartTitleElement(root: Element): Element | undefined {
  const chart = root.localName === 'chart' && root.namespaceURI === CHART_NS
    ? root
    : chartElementDescendants(root, 'chart')[0];
  if (!chart) return undefined;
  const children = chart.childNodes;
  for (let i = 0; i < children.length; i++) {
    const node = children[i];
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    const name = el.localName || (el.nodeName ? el.nodeName.split(':').pop() : '');
    if (el.namespaceURI !== CHART_NS) continue;
    if (name === 'plotArea') return undefined;
    // CT_Chart puts the chart title directly under <c:chart>; a nested
    // <c:title> before the plot area is not a legal chart title source.
    if (name === 'title') return el;
  }
  return undefined;
}

/** Direct chart-namespace children of a DOM element. */
export function chartElementChildren(parent: Element, localName: string): Element[] {
  const out: Element[] = [];
  try {
    const children = parent.childNodes;
    for (let i = 0; i < children.length; i++) {
      const node = children[i];
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      const name = el.localName || (el.nodeName ? el.nodeName.split(':').pop() : '');
      if (name === localName && el.namespaceURI === CHART_NS) out.push(el);
    }
  } catch {
    /* best-effort */
  }
  return out;
}

/** Point texts keyed by c:pt idx; gaps stay '' so x/y can be paired by index. */
export function chartPointTextsByIndex(container: Element): string[] {
  const pts = chartElementDescendants(container, 'pt');
  const values: string[] = [];
  for (const pt of pts) {
    const idx = Number(pt.getAttribute('idx'));
    if (!Number.isInteger(idx) || idx < 0 || idx > 100000) continue;
    while (values.length <= idx) values.push('');
    values[idx] = chartElementDescendants(pt, 'v')[0]?.textContent ?? '';
  }
  return values;
}

/** Per-series scatter x point text from c:ser/c:xVal caches, in series order. */
export function chartScatterXValuesBySeries(root: Element): string[][] {
  return chartElementDescendants(root, 'ser').map((ser) => {
    const xVal = chartElementDescendants(ser, 'xVal')[0];
    if (!xVal) return [];
    return chartPointTextsByIndex(xVal);
  });
}

/** Scatter x list for a series, including when it equals the first list. */
export function scatterSeriesXList(lists: string[][], index: number): string[] | undefined {
  const list = lists[index];
  return list && list.length > 0 ? list : undefined;
}

/**
 * Exposes a scatter x list whose series has no y container at all. Every
 * non-empty x point is unpaired, so readers report the omission instead of
 * dropping the list silently.
 */
export function scatterXOnlyValues(xText: string[]): { xValues: number[]; unpairedX: number } {
  const present = xText.filter((value) => value !== '');
  return {
    xValues: present.map((value) => parseFloat(value) || 0),
    unpairedX: present.length,
  };
}

/**
 * Pairs scatter x and y point texts by c:pt idx. A point missing either
 * coordinate is dropped instead of shifting another point's x or y; each
 * side's unpaired count is returned so readers can report discarded data.
 */
export function pairScatterPointTexts(
  xText: string[],
  yText: string[]
): { xValues: number[]; values: number[]; unpairedX: number; unpairedY: number } {
  const xValues: number[] = [];
  const values: number[] = [];
  const count = Math.min(xText.length, yText.length);
  for (let i = 0; i < count; i++) {
    if (!xText[i] || !yText[i]) continue;
    xValues.push(parseFloat(xText[i]) || 0);
    values.push(parseFloat(yText[i]) || 0);
  }
  return {
    xValues,
    values,
    unpairedX: xText.filter((value) => value !== '').length - xValues.length,
    unpairedY: yText.filter((value) => value !== '').length - values.length,
  };
}

export function findChartNamespacePlotElements(xml: string): ChartNamespaceElement[] {
  return findChartNamespaceElements(xml, PLOT_LOCAL_NAMES_ALL);
}

export function hasPlotLocalName(xml: string, localName: string): boolean {
  return findChartNamespaceElements(xml, [localName]).length > 0;
}

/** Every plot element occurrence, in stable tag-list order (duplicates kept). */
export function findPlotElementLocalNames(xml: string): string[] {
  const elements = findChartNamespacePlotElements(xml);
  const found: string[] = [];
  for (const tag of PLOT_LOCAL_NAMES_ALL) {
    for (const element of elements) {
      if (element.localName === tag) found.push(tag);
    }
  }
  return found;
}

/** Reverse lookup derived from the writer's variant table — one matrix only. */
const BAR_2D_TYPE_BY_DIR_AND_GROUPING = Object.fromEntries(
  Object.entries(BAR_2D_VARIANTS).map(([type, v]) => [`${v.barDir}:${v.grouping}`, type])
) as Record<string, string>;

/**
 * Reads the c:barChart family name (direction + grouping) from a chart part.
 * Scoped to the first <c:barChart> block so a combo's line series or a 3D bar
 * part can never contribute its own c:grouping. Unknown/absent grouping is
 * clustered, matching the OOXML default.
 */
export function detectBar2DChartTypeFromXml(chartXml: string): string | undefined {
  const bar = findChartNamespaceElements(chartXml, ['barChart'])[0];
  if (!bar) return undefined;
  const gt = tagEnd(chartXml, bar.start);
  let block = chartXml.slice(bar.start);
  if (gt >= 0) {
    if (bar.selfClosing) {
      block = chartXml.slice(bar.start, gt + 1);
    } else {
      const close = chartXml.indexOf(`</${bar.qualifiedName}>`, gt + 1);
      if (close >= 0) block = chartXml.slice(bar.start, close + bar.qualifiedName.length + 3);
    }
  }
  const dir = new RegExp(`<${ANY_PREFIX}barDir[^>]*val="bar"`).test(block) ? 'bar' : 'col';
  const groupingVal = new RegExp(`<${ANY_PREFIX}grouping[^>]*val="([^"]+)"`).exec(block)?.[1];
  const grouping =
    groupingVal === 'stacked' || groupingVal === 'percentStacked' ? groupingVal : 'clustered';
  return BAR_2D_TYPE_BY_DIR_AND_GROUPING[`${dir}:${grouping}`];
}

export function detectChartKindFromXml(chartXml: string): { is3D: boolean; chartType: string } {
  const plots = findChartNamespacePlotElements(chartXml);
  const has = (localName: string): boolean => plots.some((plot) => plot.localName === localName);
  // Actual plot elements win; <c:view3D> is only a fallback hint for
  // otherwise-unmodeled 3D parts, so a valid 2D plot carrying view3D stays 2D.
  if (has('doughnutChart')) return { is3D: false, chartType: 'doughnut' };
  if (has('pieChart')) return { is3D: false, chartType: 'pie' };
  // Supported combo: exactly one barChart + one lineChart. The update gate
  // validates the axis signature before it allows a merge; read/export may
  // still surface a foreign two-plot part as combo without making it writable.
  if (plots.length >= 2 && plots.length <= 3 && has('barChart') && has('lineChart')
    && plots.every((plot) => plot.localName === 'barChart' || plot.localName === 'lineChart')) {
    return { is3D: false, chartType: 'combo' };
  }
  if (has('lineChart')) return { is3D: false, chartType: 'line' };
  if (has('areaChart')) return { is3D: false, chartType: 'area' };
  const bar2DType = detectBar2DChartTypeFromXml(chartXml);
  if (bar2DType) return { is3D: false, chartType: bar2DType };
  // Scatter is checked after the families that predate it so a legal combo
  // part (e.g. bar+scatter) keeps its previous family for read/export; pure
  // scatter parts still resolve here.
  if (has('scatterChart')) return { is3D: false, chartType: 'scatter' };
  const allChartElements = findChartNamespaceElements(chartXml);
  const has3DPlot = PLOT_LOCAL_NAMES_3D.some((tag) => has(tag))
    || allChartElements.some((element) => /3DChart$/.test(element.localName))
    || allChartElements.some((element) => element.localName === 'view3D');
  if (has3DPlot) {
    if (has('bar3DChart')) {
      const barDir = new RegExp(`<${ANY_PREFIX}barDir[^>]*val="(bar|col)"`).exec(chartXml)?.[1];
      return { is3D: true, chartType: barDir === 'bar' ? 'bar3D' : 'column3D' };
    }
    if (has('line3DChart')) return { is3D: true, chartType: 'line3D' };
    if (has('pie3DChart')) return { is3D: true, chartType: 'pie3D' };
    return { is3D: true, chartType: 'column3D' };
  }
  return { is3D: false, chartType: 'column' };
}

// NOTE: DocxParser-backed factories (createEmptyDocument,
// createBlankDocxBuffer) live in ../parser.ts — importing DocxParser here
// would reintroduce an ESM value cycle (parser <-> factory).
