/**
 * Structural chart updates that keep the existing chart part (0.5.3, DN1).
 *
 *   Case A  addSeriesInPlace      append series built from the last series of the plot
 *   Case B  removeSeriesInPlace   drop series, renumber, keep legend entries consistent
 *   Case C  convertBarFamilyInPlace  column <-> bar, clustered <-> stacked <-> percent
 *   Case D  describeLoss          what a rebuild would not keep (used by the refusal and the report)
 *
 * Everything here is text-level on the chart part so nodes the engine does not
 * model (per-point colours, error bars, extensions, theme links) stay byte-identical.
 * A shape that cannot be changed safely throws E_VALIDATION with a remedy and
 * writes nothing; the caller decides whether `allowRebuild` offers a way out.
 */
import { chartNamespaceScopeFromXml, findChartNamespaceElementsScoped } from '../parser/factory';
import {
  attrValue,
  chartQualifiedName,
  escapeXml,
  findChartElement,
  setAttribute,
  tagEnd,
  type ChartNamespaceScope,
  type Span,
} from './chart-xml-spans';

export class ChartStructureRefusal extends Error {
  readonly code = 'E_VALIDATION';
  readonly retryable = false;
  constructor(message: string, readonly remedy: string) {
    super(message);
    this.name = 'ChartStructureRefusal';
  }
}

function refuse(chartPath: string, detail: string, remedy: string): ChartStructureRefusal {
  return new ChartStructureRefusal(`E_VALIDATION: chart update failed: ${detail} (${chartPath}).`, remedy);
}

const REBUILD_TAIL = 'To do it anyway pass allowRebuild:true; the chart part is then regenerated and formatting ASTra does not model is lost.';

interface Child { local: string; start: number; end: number }

function childrenOf(parentXml: string, scope: ChartNamespaceScope): Child[] {
  return findChartNamespaceElementsScoped(parentXml, scope)
    .filter((e) => e.depth === 1 && e.end >= 0)
    .map((e) => ({ local: e.localName, start: e.start, end: e.end }));
}

const PLOT_LOCALS = new Set([
  'barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart', 'area3DChart', 'pieChart', 'pie3DChart',
  'doughnutChart', 'scatterChart', 'bubbleChart', 'stockChart', 'surfaceChart', 'surface3DChart', 'ofPieChart', 'radarChart',
]);
const AXIS_LOCALS = new Set(['catAx', 'valAx', 'dateAx', 'serAx']);
const SERIES_PLOTS = new Set(['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'areaChart']);

interface PlotArea { span: Span; xml: string }

function plotAreaOf(xml: string, scope: ChartNamespaceScope, chartPath: string): PlotArea {
  const span = findChartElement(xml, 'plotArea', scope);
  if (!span) throw refuse(chartPath, 'the chart part has no plot area', 'Remedy: delete the chart and insert it again.');
  return { span, xml: xml.slice(span.start, span.end) };
}

function absolute(area: PlotArea, child: Child): Span {
  return { start: area.span.start + child.start, end: area.span.start + child.end };
}

function splice(xml: string, span: Span, text: string): string {
  return xml.slice(0, span.start) + text + xml.slice(span.end);
}

// ---------------------------------------------------------------------------
// Case C: type change inside the column / bar family
// ---------------------------------------------------------------------------

export const BAR_FAMILY_TYPES = [
  'column', 'bar', 'columnStacked', 'barStacked', 'columnPercentStacked', 'barPercentStacked',
] as const;
export type BarFamilyType = (typeof BAR_FAMILY_TYPES)[number];

export function isBarFamilyType(type: string): type is BarFamilyType {
  return (BAR_FAMILY_TYPES as readonly string[]).includes(type);
}

function barVariantOf(type: BarFamilyType): { barDir: 'col' | 'bar'; grouping: 'clustered' | 'stacked' | 'percentStacked' } {
  const barDir = type.startsWith('bar') ? 'bar' : 'col';
  const grouping = type.endsWith('PercentStacked') ? 'percentStacked' : type.endsWith('Stacked') ? 'stacked' : 'clustered';
  return { barDir, grouping };
}

function editChildAttr(
  xml: string,
  scope: ChartNamespaceScope,
  containerLocals: ReadonlySet<string>,
  childLocal: string,
  attribute: string,
  value: string,
  chartPath: string,
  within: 'plot' | 'axis',
  axisLocal?: string
): string {
  const area = plotAreaOf(xml, scope, chartPath);
  const container = childrenOf(area.xml, scope).find(
    (c) => containerLocals.has(c.local) && (within === 'plot' || c.local === axisLocal)
  );
  if (!container) return xml;
  const span = absolute(area, container);
  const containerXml = xml.slice(span.start, span.end);
  const child = childrenOf(containerXml, scope).find((c) => c.local === childLocal);
  if (!child) return xml;
  const patched = setAttribute(containerXml.slice(child.start, child.end), attribute, value);
  return splice(xml, span, containerXml.slice(0, child.start) + patched + containerXml.slice(child.end));
}

/**
 * Converts a column/bar chart to another member of the same family. Only what the
 * conversion needs is changed: barDir, axis positions, grouping, overlap (100 for
 * stacked forms) and the percent number format when it is still the default.
 */
export function convertBarFamilyInPlace(
  chartXml: string,
  chartPath: string,
  fromType: BarFamilyType,
  toType: BarFamilyType
): { xml: string; styleNotApplied: string[] } {
  let xml = chartXml;
  const scope = chartNamespaceScopeFromXml(xml);
  const notApplied: string[] = [];
  const from = barVariantOf(fromType);
  const to = barVariantOf(toType);

  const area = plotAreaOf(xml, scope, chartPath);
  const top = childrenOf(area.xml, scope);
  const plots = top.filter((c) => PLOT_LOCALS.has(c.local));
  const axes = top.filter((c) => AXIS_LOCALS.has(c.local));
  if (plots.length !== 1 || plots[0].local !== 'barChart') {
    throw refuse(chartPath, `the chart has ${plots.length} plots (${plots.map((p) => p.local).join(', ') || 'none'}), so its type cannot be changed in place`, `Remedy: change the type of a single-plot column/bar chart. ${REBUILD_TAIL}`);
  }
  const plotXml = area.xml.slice(plots[0].start, plots[0].end);
  const plotKids = childrenOf(plotXml, scope);
  const plotAxIds = plotKids.filter((c) => c.local === 'axId').map((c) => attrValue(plotXml.slice(c.start, c.end), 'val'));
  const catAxes = axes.filter((a) => a.local === 'catAx');
  const valAxes = axes.filter((a) => a.local === 'valAx');
  const axisId = (a: Child): string | undefined => {
    const axXml = area.xml.slice(a.start, a.end);
    const id = childrenOf(axXml, scope).find((c) => c.local === 'axId');
    return id ? attrValue(axXml.slice(id.start, id.end), 'val') : undefined;
  };
  const catId = catAxes.length === 1 ? axisId(catAxes[0]) : undefined;
  const valId = valAxes.length === 1 ? axisId(valAxes[0]) : undefined;
  const idsOk = axes.length === 2 && catId !== undefined && valId !== undefined && catId !== valId
    && plotAxIds.length === 2 && new Set(plotAxIds).size === 2 && plotAxIds.includes(catId) && plotAxIds.includes(valId);
  if (!idsOk) {
    throw refuse(
      chartPath,
      'the plot does not use exactly one category axis and one value axis of its own (secondary axes, shared axes or missing axis ids)',
      `Remedy: change the type of a chart with one category axis and one value axis. ${REBUILD_TAIL}`
    );
  }
  const plotKidLocals = new Set(plotKids.map((c) => c.local));
  if (!plotKidLocals.has('barDir') || !plotKidLocals.has('grouping')) {
    throw refuse(chartPath, 'the bar plot has no barDir or grouping element', 'Remedy: delete the chart and insert it again.');
  }

  if (from.barDir !== to.barDir) {
    xml = editChildAttr(xml, scope, PLOT_LOCALS, 'barDir', 'val', to.barDir, chartPath, 'plot');
    const catPos = to.barDir === 'bar' ? 'l' : 'b';
    const valPos = to.barDir === 'bar' ? 'b' : 'l';
    xml = editChildAttr(xml, scope, AXIS_LOCALS, 'axPos', 'val', catPos, chartPath, 'axis', 'catAx');
    xml = editChildAttr(xml, scope, AXIS_LOCALS, 'axPos', 'val', valPos, chartPath, 'axis', 'valAx');
  }
  if (from.grouping !== to.grouping) {
    xml = editChildAttr(xml, scope, PLOT_LOCALS, 'grouping', 'val', to.grouping, chartPath, 'plot');
    xml = adjustOverlap(xml, scope, chartPath, to.grouping !== 'clustered');
    xml = adjustPercentFormat(xml, scope, chartPath, from.grouping, to.grouping, notApplied);
  }
  return { xml, styleNotApplied: notApplied };
}

function adjustOverlap(xml: string, scope: ChartNamespaceScope, chartPath: string, stacked: boolean): string {
  const area = plotAreaOf(xml, scope, chartPath);
  const plot = childrenOf(area.xml, scope).find((c) => PLOT_LOCALS.has(c.local));
  if (!plot) return xml;
  const span = absolute(area, plot);
  const plotXml = xml.slice(span.start, span.end);
  const kids = childrenOf(plotXml, scope);
  const overlap = kids.find((c) => c.local === 'overlap');
  if (stacked) {
    if (overlap) {
      const current = attrValue(plotXml.slice(overlap.start, overlap.end), 'val');
      if (current === '100') return xml;
      const patched = setAttribute(plotXml.slice(overlap.start, overlap.end), 'val', '100');
      return splice(xml, span, plotXml.slice(0, overlap.start) + patched + plotXml.slice(overlap.end));
    }
    const before = kids.find((c) => c.local === 'serLines') ?? kids.find((c) => c.local === 'axId');
    const at = before ? before.start : plotXml.lastIndexOf('</');
    const tag = chartQualifiedName(plotXml, 'overlap');
    return splice(xml, span, `${plotXml.slice(0, at)}<${tag} val="100"/>${plotXml.slice(at)}`);
  }
  if (overlap && attrValue(plotXml.slice(overlap.start, overlap.end), 'val') === '100') {
    return splice(xml, span, plotXml.slice(0, overlap.start) + plotXml.slice(overlap.end));
  }
  return xml;
}

function adjustPercentFormat(
  xml: string,
  scope: ChartNamespaceScope,
  chartPath: string,
  fromGrouping: string,
  toGrouping: string,
  notApplied: string[]
): string {
  const enteringPercent = toGrouping === 'percentStacked' && fromGrouping !== 'percentStacked';
  const leavingPercent = fromGrouping === 'percentStacked' && toGrouping !== 'percentStacked';
  if (!enteringPercent && !leavingPercent) return xml;
  const area = plotAreaOf(xml, scope, chartPath);
  const valAx = childrenOf(area.xml, scope).find((c) => c.local === 'valAx');
  if (!valAx) return xml;
  const span = absolute(area, valAx);
  const axXml = xml.slice(span.start, span.end);
  const numFmt = childrenOf(axXml, scope).find((c) => c.local === 'numFmt');
  const fmtXml = numFmt ? axXml.slice(numFmt.start, numFmt.end) : undefined;
  const code = fmtXml ? attrValue(fmtXml, 'formatCode') : undefined;
  const linked = fmtXml ? attrValue(fmtXml, 'sourceLinked') : undefined;
  const isDefault = !fmtXml || linked === '1' || code === 'General';
  if (enteringPercent) {
    if (!isDefault) {
      notApplied.push(`value axis number format "${code}" kept; a percent chart usually shows 0%`);
      return xml;
    }
    if (fmtXml) {
      const patched = setAttribute(setAttribute(fmtXml, 'formatCode', '0%'), 'sourceLinked', '0');
      return splice(xml, span, axXml.slice(0, numFmt!.start) + patched + axXml.slice(numFmt!.end));
    }
    return xml;
  }
  // Leaving percent: undo only the format the percent form itself wrote.
  if (fmtXml && code === '0%' && linked === '0') {
    const patched = setAttribute(setAttribute(fmtXml, 'formatCode', 'General'), 'sourceLinked', '1');
    return splice(xml, span, axXml.slice(0, numFmt!.start) + patched + axXml.slice(numFmt!.end));
  }
  if (fmtXml && !isDefault) notApplied.push(`value axis number format "${code}" kept after leaving the percent form`);
  return xml;
}

// ---------------------------------------------------------------------------
// Shared: series access for cases A and B
// ---------------------------------------------------------------------------

interface SeriesPlot {
  plotSpan: Span;
  plotXml: string;
  series: Span[]; // absolute, in document order
}

function singleSeriesPlot(xml: string, scope: ChartNamespaceScope, chartPath: string, what: string): SeriesPlot {
  const area = plotAreaOf(xml, scope, chartPath);
  const plots = childrenOf(area.xml, scope).filter((c) => PLOT_LOCALS.has(c.local));
  if (plots.length !== 1 || !SERIES_PLOTS.has(plots[0].local)) {
    throw refuse(
      chartPath,
      `${what} works on one column, bar, line or area plot; this chart has ${plots.length === 1 ? plots[0].local : plots.length + ' plots'}`,
      `Remedy: ${REBUILD_TAIL}`
    );
  }
  const plotSpan = absolute(area, plots[0]);
  const plotXml = xml.slice(plotSpan.start, plotSpan.end);
  const series = childrenOf(plotXml, scope)
    .filter((c) => c.local === 'ser')
    .map((c) => ({ start: plotSpan.start + c.start, end: plotSpan.start + c.end }));
  return { plotSpan, plotXml, series };
}

/** True when the chart has one plot that case A and B can change in place. */
export function supportsInPlaceSeriesChange(chartXml: string): boolean {
  try {
    const scope = chartNamespaceScopeFromXml(chartXml);
    singleSeriesPlot(chartXml, scope, 'chart part', 'series change');
    return true;
  } catch {
    return false;
  }
}

function seriesNumber(serXml: string, scope: ChartNamespaceScope, local: 'idx' | 'order'): number {
  const child = childrenOf(serXml, scope).find((c) => c.local === local);
  const value = child ? Number(attrValue(serXml.slice(child.start, child.end), 'val')) : Number.NaN;
  return Number.isInteger(value) ? value : Number.NaN;
}

function columnLetter(index: number): string {
  let letter = '';
  let n = index;
  while (n >= 0) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

// ---------------------------------------------------------------------------
// Case A: add series
// ---------------------------------------------------------------------------

export interface AddedSeries {
  name: string;
  values: number[];
  /** Six hex digits, no '#'. */
  color: string;
}

const COPIED_CHILDREN = new Set(['invertIfNegative', 'shape', 'smooth', 'explosion']);
const DROPPED_CHILDREN = new Set(['dPt', 'trendline', 'errBars', 'extLst']);
const KNOWN_CHILDREN = new Set([
  'idx', 'order', 'tx', 'spPr', 'marker', 'dLbls', 'cat', 'val',
  ...COPIED_CHILDREN, ...DROPPED_CHILDREN,
]);

function solidFill(color: string): string {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

function buildAddedSeries(
  templateXml: string,
  scope: ChartNamespaceScope,
  chartPath: string,
  plotLocal: string,
  next: { idx: number; order: number; column: number },
  added: AddedSeries
): string {
  const kids = childrenOf(templateXml, scope);
  const unknown = kids.find((c) => !KNOWN_CHILDREN.has(c.local));
  if (unknown) {
    throw refuse(
      chartPath,
      `the series used as the template carries <c:${unknown.local}>, which ASTra does not know how to carry to a new series`,
      `Remedy: ${REBUILD_TAIL}`
    );
  }
  const open = templateXml.slice(0, tagEnd(templateXml, 0) + 1);
  const closeStart = templateXml.lastIndexOf('</');
  const close = templateXml.slice(closeStart);
  const q = (local: string): string => chartQualifiedName(templateXml, local);
  const lineSeries = plotLocal === 'lineChart' || plotLocal === 'line3DChart';
  const col = columnLetter(next.column);
  const last = added.values.length + 1;
  const points = added.values.map((v, i) => `<${q('pt')} idx="${i}"><${q('v')}>${v}</${q('v')}></${q('pt')}>`).join('');
  let out = open;
  for (const child of kids) {
    const text = templateXml.slice(child.start, child.end);
    switch (child.local) {
      case 'idx':
        out += `<${q('idx')} val="${next.idx}"/>`;
        break;
      case 'order':
        out += `<${q('order')} val="${next.order}"/>`;
        break;
      case 'tx':
        out += `<${q('tx')}><${q('strRef')}><${q('f')}>Sheet1!$${col}$1</${q('f')}><${q('strCache')}>` +
          `<${q('ptCount')} val="1"/><${q('pt')} idx="0"><${q('v')}>${escapeXml(added.name)}</${q('v')}></${q('pt')}>` +
          `</${q('strCache')}></${q('strRef')}></${q('tx')}>`;
        break;
      case 'spPr':
        out += lineSeries
          ? `<${q('spPr')}><a:ln w="28575" cap="rnd">${solidFill(added.color)}<a:round/></a:ln></${q('spPr')}>`
          : `<${q('spPr')}>${solidFill(added.color)}</${q('spPr')}>`;
        break;
      case 'marker':
        // Same marker shape, own colour: every fill in the marker follows the new series colour.
        out += text.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, solidFill(added.color));
        break;
      case 'dLbls': {
        // Series-wide label settings carry over; labels of single points do not.
        const lblKids = childrenOf(text, scope).filter((c) => c.local === 'dLbl');
        let kept = text;
        for (const lbl of [...lblKids].reverse()) kept = kept.slice(0, lbl.start) + kept.slice(lbl.end);
        out += kept;
        break;
      }
      case 'cat':
        out += text;
        break;
      case 'val':
        out += `<${q('val')}><${q('numRef')}><${q('f')}>Sheet1!$${col}$2:$${col}$${last}</${q('f')}>` +
          `<${q('numCache')}><${q('formatCode')}>General</${q('formatCode')}><${q('ptCount')} val="${added.values.length}"/>${points}</${q('numCache')}>` +
          `</${q('numRef')}></${q('val')}>`;
        break;
      default:
        if (COPIED_CHILDREN.has(child.local)) out += text;
        // dPt, trendline, errBars, extLst: never carried to a new series.
        break;
    }
  }
  if (!kids.some((c) => c.local === 'spPr')) {
    // A template without its own colour still gets one on the new series (schema order: after tx).
    const txEnd = out.indexOf(`</${q('tx')}>`);
    const at = txEnd >= 0 ? txEnd + `</${q('tx')}>`.length : open.length;
    const spPr = lineSeries
      ? `<${q('spPr')}><a:ln w="28575" cap="rnd">${solidFill(added.color)}<a:round/></a:ln></${q('spPr')}>`
      : `<${q('spPr')}>${solidFill(added.color)}</${q('spPr')}>`;
    out = out.slice(0, at) + spPr + out.slice(at);
  }
  return out + close;
}

export function addSeriesInPlace(
  chartXml: string,
  chartPath: string,
  added: AddedSeries[],
  existingCount: number
): string {
  let xml = chartXml;
  const scope = chartNamespaceScopeFromXml(xml);
  const plot = singleSeriesPlot(xml, scope, chartPath, 'adding a series');
  if (plot.series.length === 0) {
    throw refuse(chartPath, 'the plot has no series to use as a template', 'Remedy: delete the chart and insert it again with data.');
  }
  if (plot.series.length !== existingCount) {
    throw refuse(chartPath, `the part carries ${plot.series.length} series but the chart reads ${existingCount}`, 'Remedy: re-read the chart with inspect_document and retry.');
  }
  const plotLocal = findPlotLocal(xml, scope, chartPath);
  const all = findChartSeriesAll(xml, scope);
  let nextIdx = Math.max(-1, ...all.map((s) => seriesNumber(xml.slice(s.start, s.end), scope, 'idx')).filter(Number.isFinite)) + 1;
  let nextOrder = Math.max(-1, ...all.map((s) => seriesNumber(xml.slice(s.start, s.end), scope, 'order')).filter(Number.isFinite)) + 1;
  const template = xml.slice(plot.series[plot.series.length - 1].start, plot.series[plot.series.length - 1].end);
  let built = '';
  let column = existingCount + 1;
  for (const entry of added) {
    built += buildAddedSeries(template, scope, chartPath, plotLocal, { idx: nextIdx++, order: nextOrder++, column: column++ }, entry);
  }
  const insertAt = plot.series[plot.series.length - 1].end;
  xml = xml.slice(0, insertAt) + built + xml.slice(insertAt);
  return xml;
}

function findPlotLocal(xml: string, scope: ChartNamespaceScope, chartPath: string): string {
  const area = plotAreaOf(xml, scope, chartPath);
  return childrenOf(area.xml, scope).find((c) => PLOT_LOCALS.has(c.local))?.local ?? '';
}

function findChartSeriesAll(xml: string, scope: ChartNamespaceScope): Span[] {
  const area = plotAreaOf(xml, scope, 'chart part');
  const result: Span[] = [];
  for (const plot of childrenOf(area.xml, scope).filter((c) => PLOT_LOCALS.has(c.local))) {
    const plotXml = area.xml.slice(plot.start, plot.end);
    for (const ser of childrenOf(plotXml, scope).filter((c) => c.local === 'ser')) {
      result.push({ start: area.span.start + plot.start + ser.start, end: area.span.start + plot.start + ser.end });
    }
  }
  return result;
}

// ---------------------------------------------------------------------------
// Case B: remove series
// ---------------------------------------------------------------------------

/** Maps names and one-based positions to zero-based indexes; every problem lists the chart's series. */
export function resolveSeriesRemovals(
  request: ReadonlyArray<string | number>,
  existingNames: readonly string[],
  chartPath: string
): number[] {
  const list = existingNames.map((n, i) => `${i + 1}: "${n}"`).join(', ');
  const picked = new Set<number>();
  for (const item of request) {
    let index: number;
    if (typeof item === 'number') {
      if (!Number.isInteger(item) || item < 1 || item > existingNames.length) {
        throw refuse(chartPath, `removeSeries position ${String(item)} is outside 1..${existingNames.length}`, `Remedy: the chart's series are ${list}.`);
      }
      index = item - 1;
    } else {
      const matches = existingNames.flatMap((n, i) => (n === item ? [i] : []));
      if (matches.length === 0) {
        throw refuse(chartPath, `removeSeries name "${item}" matches no series`, `Remedy: the chart's series are ${list}.`);
      }
      if (matches.length > 1) {
        throw refuse(chartPath, `removeSeries name "${item}" matches ${matches.length} series`, `Remedy: remove by one-based position instead. The chart's series are ${list}.`);
      }
      index = matches[0];
    }
    if (picked.has(index)) {
      throw refuse(chartPath, `removeSeries names series ${index + 1} twice`, `Remedy: list each series once. The chart's series are ${list}.`);
    }
    picked.add(index);
  }
  if (picked.size >= existingNames.length) {
    throw refuse(chartPath, 'removeSeries would remove every series', 'Remedy: keep at least one series, or delete the chart.');
  }
  return [...picked].sort((a, b) => a - b);
}

export function removeSeriesInPlace(
  chartXml: string,
  chartPath: string,
  removeIndexes: readonly number[],
  expectedCount: number
): string {
  let xml = chartXml;
  const scope = chartNamespaceScopeFromXml(xml);
  const plot = singleSeriesPlot(xml, scope, chartPath, 'removing a series');
  if (plot.series.length !== expectedCount) {
    throw refuse(chartPath, `the part carries ${plot.series.length} series but the chart reads ${expectedCount}`, 'Remedy: re-read the chart with inspect_document and retry.');
  }
  const drop = new Set(removeIndexes);
  const meta = plot.series.map((s) => {
    const serXml = xml.slice(s.start, s.end);
    return { span: s, idx: seriesNumber(serXml, scope, 'idx'), order: seriesNumber(serXml, scope, 'order') };
  });
  if (meta.some((m) => !Number.isFinite(m.idx) || !Number.isFinite(m.order))) {
    throw refuse(chartPath, 'a series has no numeric idx or order', 'Remedy: delete the chart and insert it again.');
  }
  const kept = meta.filter((_, i) => !drop.has(i));
  const rank = (values: number[]): Map<number, number> =>
    new Map([...values].sort((a, b) => a - b).map((value, position) => [value, position]));
  const idxRank = rank(kept.map((m) => m.idx));
  const orderRank = rank(kept.map((m) => m.order));
  const idxMap = new Map<number, number | null>();
  meta.forEach((m, i) => idxMap.set(m.idx, drop.has(i) ? null : (idxRank.get(m.idx) as number)));

  // Edit from the end of the part to the start so earlier offsets stay valid.
  const edits: Array<{ span: Span; text: string }> = [];
  meta.forEach((m, i) => {
    if (drop.has(i)) {
      edits.push({ span: m.span, text: '' });
      return;
    }
    let serXml = xml.slice(m.span.start, m.span.end);
    const set = (local: 'idx' | 'order', value: number): void => {
      const child = childrenOf(serXml, scope).find((c) => c.local === local);
      if (!child) return;
      serXml = serXml.slice(0, child.start) + setAttribute(serXml.slice(child.start, child.end), 'val', String(value)) + serXml.slice(child.end);
    };
    set('idx', idxRank.get(m.idx) as number);
    set('order', orderRank.get(m.order) as number);
    edits.push({ span: m.span, text: serXml });
  });
  for (const edit of edits.sort((a, b) => b.span.start - a.span.start)) {
    xml = splice(xml, edit.span, edit.text);
  }
  return fixLegendEntries(xml, scope, idxMap, chartPath);
}

function fixLegendEntries(
  chartXml: string,
  scope: ChartNamespaceScope,
  idxMap: Map<number, number | null>,
  chartPath: string
): string {
  let xml = chartXml;
  const legend = findChartElement(xml, 'legend', scope);
  if (!legend) return xml;
  const legendXml = xml.slice(legend.start, legend.end);
  const entries = childrenOf(legendXml, scope).filter((c) => c.local === 'legendEntry');
  if (entries.length === 0) return xml;
  let next = legendXml;
  for (const entry of [...entries].reverse()) {
    const entryXml = legendXml.slice(entry.start, entry.end);
    const idxChild = childrenOf(entryXml, scope).find((c) => c.local === 'idx');
    const idx = idxChild ? Number(attrValue(entryXml.slice(idxChild.start, idxChild.end), 'val')) : Number.NaN;
    if (!idxMap.has(idx)) {
      throw refuse(chartPath, `a legend entry points at index ${String(idx)}, which no series carries`, 'Remedy: pass allowRebuild:true to regenerate the chart.');
    }
    const mapped = idxMap.get(idx);
    if (mapped === null || mapped === undefined) {
      next = next.slice(0, entry.start) + next.slice(entry.end);
    } else if (mapped !== idx && idxChild) {
      const patched = entryXml.slice(0, idxChild.start) + setAttribute(entryXml.slice(idxChild.start, idxChild.end), 'val', String(mapped)) + entryXml.slice(idxChild.end);
      next = next.slice(0, entry.start) + patched + next.slice(entry.end);
    }
  }
  xml = splice(xml, legend, next);
  return xml;
}

// ---------------------------------------------------------------------------
// After a structure change: formulas follow the regenerated workbook layout
// ---------------------------------------------------------------------------

/**
 * Points every series' name, category and value formulas at the canonical
 * workbook layout (A = categories, B.. = series in order). Only formulas are
 * touched; cached values stay as they are.
 */
export function repointSeriesFormulas(chartXml: string, chartPath: string): string {
  let xml = chartXml;
  const scope = chartNamespaceScopeFromXml(xml);
  const all = findChartSeriesAll(xml, scope);
  const edits: Array<{ span: Span; text: string }> = [];
  all.forEach((span, position) => {
    let serXml = xml.slice(span.start, span.end);
    const col = columnLetter(position + 1);
    const patchFormula = (local: 'tx' | 'cat' | 'val', build: (rows: number) => string): void => {
      const child = childrenOf(serXml, scope).find((c) => c.local === local);
      if (!child) return;
      const block = serXml.slice(child.start, child.end);
      const f = findChartElement(block, 'f', scope);
      if (!f) return;
      const ptCount = findChartElement(block, 'ptCount', scope);
      const rows = ptCount ? Number(attrValue(block.slice(ptCount.start, ptCount.end), 'val')) : Number.NaN;
      const text = build(Number.isFinite(rows) && rows > 0 ? rows : 1);
      const inner = block.slice(f.start, f.end);
      const open = inner.slice(0, tagEnd(inner, 0) + 1);
      const close = inner.slice(inner.lastIndexOf('</'));
      const patched = block.slice(0, f.start) + open + escapeXml(text) + close + block.slice(f.end);
      serXml = serXml.slice(0, child.start) + patched + serXml.slice(child.end);
    };
    patchFormula('tx', () => `Sheet1!$${col}$1`);
    patchFormula('cat', (rows) => `Sheet1!$A$2:$A$${rows + 1}`);
    patchFormula('val', (rows) => `Sheet1!$${col}$2:$${col}$${rows + 1}`);
    edits.push({ span, text: serXml });
  });
  for (const edit of edits.sort((a, b) => b.span.start - a.span.start)) {
    xml = splice(xml, edit.span, edit.text);
  }
  void chartPath;
  return xml;
}

// ---------------------------------------------------------------------------
// Case D: what a rebuild does not keep
// ---------------------------------------------------------------------------

/**
 * Element types worth naming in a loss report: formatting and extras a person set.
 * Structure the generator always rewrites (plot, axis and grouping elements) is left out.
 */
const WORTH_REPORTING = new Set([
  'dPt', 'dLbl', 'dLbls', 'trendline', 'errBars', 'extLst', 'txPr', 'userShapes', 'clrMapOvr', 'style', 'roundedCorners',
  'printSettings', 'explosion', 'serLines', 'dropLines', 'hiLowLines', 'upDownBars', 'legendEntry', 'layout', 'manualLayout',
  'pictureOptions', 'floor', 'backWall', 'sideWall', 'dTable', 'AlternateContent', 'invertIfNegative', 'bubble3D',
  'majorUnit', 'minorUnit', 'minorGridlines', 'dispUnits', 'title', 'spPr', 'effectLst', 'gradFill', 'pattFill', 'blipFill',
]);

/** Names of the elements in `before` that `after` no longer contains, plus rel targets the new part cannot use. */
export function describeLoss(beforeXml: string, afterXml: string, relsXml: string | null): string[] {
  const names = (xml: string): Set<string> => {
    const out = new Set<string>();
    const re = /<([A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)(?=[\s/>])/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) {
      if (m[2] === 'xml') continue;
      out.add(m[1] ? `${m[1]}${m[2]}` : m[2]);
    }
    return out;
  };
  const strip = (name: string): string => name.replace(/^[^:]+:/, '');
  const afterNames = new Set([...names(afterXml)].map(strip));
  const lost = [...names(beforeXml)].map(strip).filter((n) => WORTH_REPORTING.has(n) && !afterNames.has(n));
  const unique = [...new Set(lost)].sort();
  if (relsXml) {
    for (const tag of relsXml.match(/<(?:[\w.-]+:)?Relationship\b[^>]*>/g) ?? []) {
      const relType = /\sType\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(tag);
      const type = /\/([A-Za-z]+)$/.exec(relType?.[1] ?? relType?.[2] ?? '')?.[1];
      if (type && /^(chartStyle|chartColorStyle|themeOverride|userShapes|image)$/.test(type) && !unique.includes(`rel:${type}`)) {
        unique.push(`rel:${type}`);
      }
    }
  }
  return unique;
}
