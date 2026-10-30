import type { Bar3DShape, View3DProperties } from '../manipulators/chart3d';
import {
  chartNamespaceScopeFromXml,
  findChartNamespaceElements,
} from '../parser/factory';
import {
  attrValue,
  chartPrefix,
  chartQualifiedName,
  chartSpans,
  findChartDirectChild,
  findDirectChartTitleSpan,
  findChartElement,
  findDirectChild,
  findAllElements,
  setAttribute,
  setElementText,
  tagEnd,
  tagNameOf,
  type ChartNamespaceScope,
  type Span,
} from './chart-xml-spans';

/**
 * Preserve-and-merge chart updates: patches the data nodes an update actually
 * changes (title text, category points, series names/values/colors, 3D view
 * and bar shape) inside the existing chart part instead of regenerating it.
 * Every node the read path does not model — axes configuration, tick marks,
 * number formats, legend placement, per-point formats and fills, trendlines,
 * plot-area fills, external-data links — carries through byte-identical.
 *
 * Any shape that cannot be patched (missing point cache, non-dense point
 * indexes, numeric category axis asked for non-numeric labels, series-count or
 * chart-type skew) fails closed with E_VALIDATION + remedy; nothing is written.
 */

export interface ChartMergeExistingSeries {
  name: string;
  values: number[];
  color?: string;
  xValues?: number[];
  chartType?: 'column' | 'line';
  axis?: 'primary' | 'secondary';
}

export interface ChartMergeSeriesPatch {
  name: string;
  values: number[];
  color?: string;
  xValues?: number[];
  chartType?: 'column' | 'line';
  axis?: 'primary' | 'secondary';
}

export interface ChartMergeRequest {
  chartPath: string;
  chartType: string;
  is3D: boolean;
  title?: string;
  categories?: string[];
  series?: ChartMergeSeriesPatch[];
  view3D?: View3DProperties;
  shape?: Bar3DShape;
  holeSize?: number;
  existing: {
    title: string;
    categories: string[];
    series: ChartMergeExistingSeries[];
  };
}

const VIEW_3D_ORDER = ['rotX', 'hPercent', 'rotY', 'depthPercent', 'rAngAx', 'perspective'];

const REINDEX_REMEDY =
  're-read the chart via inspect_document; if the chart was authored elsewhere, delete it and re-insert it with action:\'chart\', operation:\'insert\', then apply the data update.';

function mergeRefusal(chartPath: string, detail: string, remedy: string): Error {
  return Object.assign(
    new Error(`E_VALIDATION: chart update failed: ${detail} (${chartPath}).`),
    { code: 'E_VALIDATION', retryable: false, remedy }
  );
}

export function escapeXml(value: string): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&');
}

function sameNumbers(a: number[], b: number[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function sameStrings(a: string[], b: string[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function normalizeHex(value: string | undefined): string | undefined {
  if (typeof value !== 'string') return undefined;
  const raw = value.replace(/^#/, '').trim();
  return raw.length > 0 ? raw.toUpperCase() : undefined;
}

function isNumericText(value: string): boolean {
  return /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(value.trim());
}

function chartColLetter(colIdx: number): string {
  let letter = '';
  let temp = colIdx;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function clampNumber(value: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, value));
}

/**
 * Rewrites the cached point list of one cache/literal container. Existing
 * points are kept byte-identical except for their `<c:v>` text (per-point
 * attributes such as formatCode survive); added points are appended with the
 * container's own spacing, removed points are dropped with their separator.
 */
function patchPointList(
  container: string,
  values: string[],
  opts: { label: string; chartPath: string; scope: ChartNamespaceScope }
): string {
  const openEnd = tagEnd(container, 0) + 1;
  const closeStart = container.lastIndexOf('</');
  if (openEnd <= 0 || closeStart < openEnd) {
    throw mergeRefusal(opts.chartPath, `${opts.label} is not a closed point container`, REINDEX_REMEDY);
  }
  const openTag = container.slice(0, openEnd);
  const closeTag = container.slice(closeStart);
  let body = container.slice(openEnd, closeStart);
  const count = findChartElement(body, 'ptCount', opts.scope);
  if (!count) {
    throw mergeRefusal(opts.chartPath, `${opts.label} has no <c:ptCount> to merge into`, REINDEX_REMEDY);
  }
  body = body.slice(0, count.start) + setAttribute(body.slice(count.start, count.end), 'val', String(values.length)) + body.slice(count.end);
  const countEnd = findChartElement(body, 'ptCount', opts.scope)!.end;
  const region = body.slice(countEnd);
  const pts = chartSpans(region, 'pt', opts.scope);
  const prefix = pts.length > 0 ? region.slice(0, pts[0].start) : region;
  const suffix = pts.length > 0 ? region.slice(pts[pts.length - 1].end) : '';
  const between = pts.length > 1 ? region.slice(pts[0].end, pts[1].start) : suffix;
  for (let i = 0; i < pts.length; i++) {
    const idx = attrValue(region.slice(pts[i].start, pts[i].end), 'idx');
    if (idx === undefined || Number(idx) !== i) {
      throw mergeRefusal(
        opts.chartPath,
        `${opts.label} carries non-dense point indexes (expected idx="${i}")`,
        REINDEX_REMEDY
      );
    }
  }
  const points: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i < pts.length) {
      const ptXml = region.slice(pts[i].start, pts[i].end);
      const v = findChartElement(ptXml, 'v', opts.scope);
      if (!v) {
        throw mergeRefusal(
          opts.chartPath,
          `${opts.label} point ${i} has no <c:v> value to merge into`,
          REINDEX_REMEDY
        );
      }
      points.push(ptXml.slice(0, v.start) + setElementText(ptXml.slice(v.start, v.end), values[i]) + ptXml.slice(v.end));
    } else {
      const pt = chartQualifiedName(container, 'pt');
      const v = chartQualifiedName(container, 'v');
      points.push(`<${pt} idx="${i}"><${v}>${escapeXml(values[i])}</${v}></${pt}>`);
    }
  }
  return openTag + body.slice(0, countEnd) + prefix + points.join(between) + suffix + closeTag;
}

function patchFormulaRange(containerXml: string, formula: string, scope: ChartNamespaceScope): string {
  const f = findChartElement(containerXml, 'f', scope);
  if (!f) return containerXml;
  return containerXml.slice(0, f.start) + setElementText(containerXml.slice(f.start, f.end), formula) + containerXml.slice(f.end);
}

function patchSeriesName(serXml: string, name: string, chartPath: string, scope: ChartNamespaceScope): string {
  const tx = findChartElement(serXml, 'tx', scope);
  if (tx) {
    const txXml = serXml.slice(tx.start, tx.end);
    const v = findChartElement(txXml, 'v', scope);
    if (!v) {
      throw mergeRefusal(chartPath, 'a series name block has no <c:v> to merge into', REINDEX_REMEDY);
    }
    const patched = txXml.slice(0, v.start) + setElementText(txXml.slice(v.start, v.end), name) + txXml.slice(v.end);
    return serXml.slice(0, tx.start) + patched + serXml.slice(tx.end);
  }
  const anchor = findChartElement(serXml, 'order', scope) ?? findChartElement(serXml, 'idx', scope);
  const at = anchor ? anchor.end : tagEnd(serXml, 0) + 1;
  const txTag = chartQualifiedName(serXml, 'tx');
  const vTag = chartQualifiedName(serXml, 'v');
  const block = `<${txTag}><${vTag}>${escapeXml(name)}</${vTag}></${txTag}>`;
  return serXml.slice(0, at) + block + serXml.slice(at);
}

function patchSeriesValues(
  serXml: string,
  values: number[],
  seriesIndex: number,
  chartPath: string,
  previousLength: number,
  scope: ChartNamespaceScope,
  valueTag: 'val' | 'yVal' = 'val'
): string {
  const val = findChartElement(serXml, valueTag, scope);
  if (!val) {
    throw mergeRefusal(chartPath, `series ${seriesIndex + 1} has no <c:${valueTag}> block to merge values into`, REINDEX_REMEDY);
  }
  const valXml = serXml.slice(val.start, val.end);
  const container = findChartElement(valXml, 'numCache', scope) ?? findChartElement(valXml, 'numLit', scope);
  if (!container) {
    throw mergeRefusal(
      chartPath,
      `series ${seriesIndex + 1} has no cached numeric values (<c:numCache>/<c:numLit>) to merge into`,
      REINDEX_REMEDY
    );
  }
  let patchedVal =
    valXml.slice(0, container.start) +
    patchPointList(valXml.slice(container.start, container.end), values.map((v) => String(v)), {
      label: `series ${seriesIndex + 1} values`,
      chartPath,
      scope,
    }) +
    valXml.slice(container.end);
  if (values.length !== previousLength) {
    const col = chartColLetter(valueTag === 'yVal' ? seriesIndex * 2 + 1 : seriesIndex + 1);
    patchedVal = patchFormulaRange(patchedVal, `Sheet1!$${col}$2:$${col}$${values.length + 1}`, scope);
  }
  return serXml.slice(0, val.start) + patchedVal + serXml.slice(val.end);
}

const SHAPE_FILL_TAGS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
const SHAPE_POST_FILL_TAGS = ['a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'];
const LINE_POST_FILL_TAGS = ['a:prstDash', 'a:custDash', 'a:round', 'a:bevel', 'a:miter', 'a:headEnd', 'a:tailEnd', 'a:extLst'];
const LINE_WIDTH_EMU = 25400;

/** Expands a self-closed element (`<x/>`) into `<x>content</x>`; null otherwise. */
function expandSelfClosedElement(elementXml: string, content: string): string | null {
  const gt = tagEnd(elementXml, 0);
  if (gt < 0 || elementXml[gt - 1] !== '/') return null;
  return `${elementXml.slice(0, gt - 1)}>${content}</${tagNameOf(elementXml)}>`;
}

/** Replaces (or inserts) the fill child of a shape/line-properties block,
 * keeping borders, effects and geometry, in schema order. */
function withFillChild(blockXml: string, color: string, postFillTags: string[]): string {
  const solid = `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
  const existing = SHAPE_FILL_TAGS.map((tag) => findDirectChild(blockXml, tag)).find((span) => span !== null);
  if (existing) return blockXml.slice(0, existing.start) + solid + blockXml.slice(existing.end);
  for (const tag of postFillTags) {
    const later = findDirectChild(blockXml, tag);
    if (later) return `${blockXml.slice(0, later.start)}${solid}${blockXml.slice(later.start)}`;
  }
  const closeStart = blockXml.lastIndexOf('</');
  if (closeStart < 0) return expandSelfClosedElement(blockXml, solid) ?? blockXml;
  return `${blockXml.slice(0, closeStart)}${solid}${blockXml.slice(closeStart)}`;
}

/** Replaces (or inserts) only the fill child of a shape-properties block, so
 * borders, effects and geometry survive an explicitly requested color. */
function withSolidFill(spPrXml: string, color: string): string {
  return withFillChild(spPrXml, color, SHAPE_POST_FILL_TAGS);
}

/** Line series carry their color in the stroke, not the shape fill. */
function withLineStroke(spPrXml: string, color: string): string {
  const ln = findDirectChild(spPrXml, 'a:ln');
  if (ln) {
    return (
      spPrXml.slice(0, ln.start) +
      withFillChild(spPrXml.slice(ln.start, ln.end), color, LINE_POST_FILL_TAGS) +
      spPrXml.slice(ln.end)
    );
  }
  const block = `<a:ln w="${LINE_WIDTH_EMU}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>`;
  for (const tag of SHAPE_POST_FILL_TAGS.slice(1)) {
    const later = findDirectChild(spPrXml, tag);
    if (later) return `${spPrXml.slice(0, later.start)}${block}${spPrXml.slice(later.start)}`;
  }
  const closeStart = spPrXml.lastIndexOf('</');
  if (closeStart < 0) return expandSelfClosedElement(spPrXml, block) ?? spPrXml;
  return `${spPrXml.slice(0, closeStart)}${block}${spPrXml.slice(closeStart)}`;
}

/** Markers drawn from the series color follow an explicit series color. */
function patchMarkerColor(serXml: string, color: string, scope: ChartNamespaceScope): string {
  const marker = findChartDirectChild(serXml, 'marker', scope);
  if (!marker) return serXml;
  const markerXml = serXml.slice(marker.start, marker.end);
  const spPr = findChartDirectChild(markerXml, 'spPr', scope);
  if (!spPr) return serXml;
  const patched =
    markerXml.slice(0, spPr.start) +
    withSolidFill(markerXml.slice(spPr.start, spPr.end), color) +
    markerXml.slice(spPr.end);
  return serXml.slice(0, marker.start) + patched + serXml.slice(marker.end);
}

function patchSeriesColor(serXml: string, color: string, lineSeries: boolean, scope: ChartNamespaceScope): string {
  const spPr = findChartDirectChild(serXml, 'spPr', scope);
  if (!spPr) {
    const tx = findChartDirectChild(serXml, 'tx', scope);
    const order = findChartDirectChild(serXml, 'order', scope);
    const idx = findChartDirectChild(serXml, 'idx', scope);
    const at = tx ? tx.end : order ? order.end : idx ? idx.end : tagEnd(serXml, 0) + 1;
    const spPrTag = chartQualifiedName(serXml, 'spPr');
    const block = lineSeries
      ? `<${spPrTag}><a:ln w="${LINE_WIDTH_EMU}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></${spPrTag}>`
      : `<${spPrTag}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></${spPrTag}>`;
    const out = serXml.slice(0, at) + block + serXml.slice(at);
    return lineSeries ? patchMarkerColor(out, color, scope) : out;
  }
  const spPrXml = serXml.slice(spPr.start, spPr.end);
  const patched = lineSeries ? withLineStroke(spPrXml, color) : withSolidFill(spPrXml, color);
  const out = serXml.slice(0, spPr.start) + patched + serXml.slice(spPr.end);
  return lineSeries ? patchMarkerColor(out, color, scope) : out;
}

function comboSeriesSlots(
  xml: string,
  spans: Span[],
  scope: ChartNamespaceScope,
  chartPath: string,
  expected: number
): number[] {
  const slots: number[] = [];
  const seen = new Set<number>();
  for (const span of spans) {
    const serXml = xml.slice(span.start, span.end);
    const idxSpan = findDirectChild(serXml, chartQualifiedName(serXml, 'idx'));
    const idx = idxSpan ? Number(attrValue(serXml.slice(idxSpan.start, idxSpan.end), 'val')) : Number.NaN;
    if (!Number.isInteger(idx) || idx < 0 || idx >= expected || seen.has(idx)) {
      throw mergeRefusal(
        chartPath,
        'the combo chart part carries non-dense or duplicate series indexes, so a positional merge cannot map values safely',
        REINDEX_REMEDY
      );
    }
    seen.add(idx);
    slots.push(idx);
  }
  return slots;
}

function patchSeriesList(
  xml: string,
  req: ChartMergeRequest,
  patches: ChartMergeSeriesPatch[],
  scope: ChartNamespaceScope
): string {
  const spans = chartSpans(xml, 'ser', scope);
  if (spans.length !== patches.length) {
    throw mergeRefusal(
      req.chartPath,
      `the chart part carries ${spans.length} series but the update supplied ${patches.length} (series-count changes need a structural rewrite)`,
      REINDEX_REMEDY
    );
  }
  const pieces: string[] = [];
  let cursor = 0;
  const combo = req.chartType === 'combo';
  const baseLineSeries = req.chartType === 'line' || req.chartType === 'line3D' || req.chartType === 'scatter';
  const valueTag: 'val' | 'yVal' = req.chartType === 'scatter' ? 'yVal' : 'val';
  const slots = combo ? comboSeriesSlots(xml, spans, scope, req.chartPath, patches.length) : spans.map((_, i) => i);
  for (let i = 0; i < spans.length; i++) {
    const slot = slots[i];
    pieces.push(xml.slice(cursor, spans[i].start));
    const existing = req.existing.series[slot];
    const patch = patches[slot];
    let serXml = xml.slice(spans[i].start, spans[i].end);
    if (!existing || existing.name !== patch.name) {
      serXml = patchSeriesName(serXml, patch.name, req.chartPath, scope);
    }
    if (!existing || !sameNumbers(existing.values, patch.values)) {
      serXml = patchSeriesValues(serXml, patch.values, slot, req.chartPath, existing ? existing.values.length : patch.values.length, scope, valueTag);
    }
    if (valueTag === 'yVal' && patch.xValues !== undefined) {
      const currentX = existing?.xValues ?? req.existing.categories.map((value) => Number(value));
      if (!sameNumbers(currentX, patch.xValues)) {
        serXml = patchScatterSeriesX(
          serXml,
          patch.xValues,
          slot,
          req.chartPath,
          scope
        );
      }
    }
    const wantColor = normalizeHex(patch.color);
    const haveColor = normalizeHex(existing?.color);
    if (wantColor !== undefined && wantColor !== haveColor) {
      const lineSeries = combo
        ? (patch.chartType ?? existing?.chartType) === 'line'
        : baseLineSeries;
      serXml = patchSeriesColor(serXml, patch.color!.replace(/^#/, ''), lineSeries, scope);
    }
    pieces.push(serXml);
    cursor = spans[i].end;
  }
  pieces.push(xml.slice(cursor));
  return pieces.join('');
}

function patchCategoryList(
  catXml: string,
  categories: string[],
  seriesIndex: number,
  chartPath: string,
  previousLength: number,
  scope: ChartNamespaceScope
): string {
  const container =
    findChartElement(catXml, 'strCache', scope) ??
    findChartElement(catXml, 'strLit', scope) ??
    findChartElement(catXml, 'numCache', scope) ??
    findChartElement(catXml, 'numLit', scope);
  if (!container) {
    throw mergeRefusal(
      chartPath,
      `series ${seriesIndex + 1} category block has no cached points (<c:strCache>/<c:strLit>/<c:numCache>/<c:numLit>) to merge into`,
      REINDEX_REMEDY
    );
  }
  const containerXml = catXml.slice(container.start, container.end);
  const numeric = /^num/.test(tagNameOf(containerXml).replace(/^[^:]+:/, ''));
  if (numeric && categories.some((c) => !isNumericText(c))) {
    throw mergeRefusal(
      chartPath,
      `series ${seriesIndex + 1} stores numeric categories; the supplied labels cannot be merged into a numeric cache`,
      'keep the existing numeric categories (omit categories), or pass numeric labels, or delete and re-insert the chart with string categories.'
    );
  }
  let patched =
    catXml.slice(0, container.start) +
    patchPointList(containerXml, categories, {
      label: `series ${seriesIndex + 1} categories`,
      chartPath,
      scope,
    }) +
    catXml.slice(container.end);
  if (categories.length !== previousLength) {
    patched = patchFormulaRange(patched, `Sheet1!$A$2:$A$${categories.length + 1}`, scope);
  }
  return patched;
}

function patchCategories(
  xml: string,
  req: ChartMergeRequest,
  categories: string[],
  scope: ChartNamespaceScope
): string {
  const spans = chartSpans(xml, 'cat', scope);
  if (spans.length === 0) {
    throw mergeRefusal(req.chartPath, 'the chart part has no category point lists to merge into', REINDEX_REMEDY);
  }
  const pieces: string[] = [];
  let cursor = 0;
  for (let i = 0; i < spans.length; i++) {
    pieces.push(xml.slice(cursor, spans[i].start));
    pieces.push(
      patchCategoryList(
        xml.slice(spans[i].start, spans[i].end),
        categories,
        i,
        req.chartPath,
        req.existing.categories.length,
        scope
      )
    );
    cursor = spans[i].end;
  }
  pieces.push(xml.slice(cursor));
  return pieces.join('');
}

/** Scatter stores x values in per-series <c:xVal> numeric lists, not <c:cat>. */
function patchScatterXValues(
  xml: string,
  req: ChartMergeRequest,
  categories: string[],
  scope: ChartNamespaceScope
): string {
  const spans = chartSpans(xml, 'xVal', scope);
  if (spans.length === 0) {
    throw mergeRefusal(
      req.chartPath,
      'the scatter chart part has no <c:xVal> lists to merge x values into',
      REINDEX_REMEDY
    );
  }
  if (spans.length !== req.existing.series.length) {
    throw mergeRefusal(
      req.chartPath,
      `the scatter chart part carries ${spans.length} x-value lists but ${req.existing.series.length} series`,
      REINDEX_REMEDY
    );
  }
  const pieces: string[] = [];
  let cursor = 0;
  let patchedCount = 0;
  for (let i = 0; i < spans.length; i++) {
    pieces.push(xml.slice(cursor, spans[i].start));
    const hasOwnX = (req.series?.[i]?.xValues ?? req.existing.series[i]?.xValues) !== undefined;
    if (hasOwnX) {
      // A series with its own x list keeps it; categories only drives series
      // that still track the shared list.
      pieces.push(xml.slice(spans[i].start, spans[i].end));
    } else {
      const xValXml = xml.slice(spans[i].start, spans[i].end);
      const container = findChartElement(xValXml, 'numCache', scope) ?? findChartElement(xValXml, 'numLit', scope);
      if (!container) {
        throw mergeRefusal(
          req.chartPath,
          `series ${i + 1} has no cached numeric x values (<c:numCache>/<c:numLit>) to merge into`,
          REINDEX_REMEDY
        );
      }
      pieces.push(
        xValXml.slice(0, container.start) +
          patchPointList(xValXml.slice(container.start, container.end), categories, {
            label: `series ${i + 1} x values`,
            chartPath: req.chartPath,
            scope,
          }) +
          xValXml.slice(container.end)
      );
      patchedCount++;
    }
    cursor = spans[i].end;
  }
  pieces.push(xml.slice(cursor));
  if (patchedCount === 0) {
    throw mergeRefusal(
      req.chartPath,
      'every scatter series stores its own x-value list, so a shared categories update would not change the chart',
      'omit categories and pass xValues inside the series entries to move each series x list.'
    );
  }
  return pieces.join('');
}

/** Rewrites one series' <c:xVal> points (the per-series x override path). */
function patchScatterSeriesX(
  serXml: string,
  xValues: number[],
  seriesIndex: number,
  chartPath: string,
  scope: ChartNamespaceScope
): string {
  const xVal = findChartElement(serXml, 'xVal', scope);
  if (!xVal) {
    throw mergeRefusal(
      chartPath,
      `series ${seriesIndex + 1} has no <c:xVal> block for its per-series x values`,
      'omit categories and pass xValues inside the series entries, or delete the chart and re-insert it with numeric x values; this part can only store x in c:xVal lists.'
    );
  }
  const xValXml = serXml.slice(xVal.start, xVal.end);
  const container = findChartElement(xValXml, 'numCache', scope) ?? findChartElement(xValXml, 'numLit', scope);
  if (!container) {
    throw mergeRefusal(
      chartPath,
      `series ${seriesIndex + 1} has no cached numeric x values (<c:numCache>/<c:numLit>) to merge into`,
      REINDEX_REMEDY
    );
  }
  const patched =
    xValXml.slice(0, container.start) +
    patchPointList(xValXml.slice(container.start, container.end), xValues.map((value) => String(value)), {
      label: `series ${seriesIndex + 1} x values`,
      chartPath,
      scope,
    }) +
    xValXml.slice(container.end);
  return serXml.slice(0, xVal.start) + patched + serXml.slice(xVal.end);
}

/** Final x list length for a series after all requested patches. */
function finalScatterXLength(req: ChartMergeRequest, seriesIndex: number): number {
  const patch = req.series?.[seriesIndex];
  if (patch?.xValues !== undefined) return patch.xValues.length;
  const existingX = req.existing.series[seriesIndex]?.xValues;
  if (existingX !== undefined) return existingX.length;
  if (req.categories !== undefined) return req.categories.length;
  return req.existing.categories.length;
}

/**
 * When a scatter update supplies data, the workbook is regenerated in the
 * canonical pair layout (A=x1, B=y1, C=x2, D=y2, ...), so every x/y formula is
 * re-pointed at its own pair. Data-free edits (title only) skip both steps and
 * keep formulas and workbook byte-identical.
 */
function patchScatterFormulas(xml: string, req: ChartMergeRequest, scope: ChartNamespaceScope): string {
  const spans = chartSpans(xml, 'ser', scope);
  if (spans.length === 0) return xml;
  const pieces: string[] = [];
  let cursor = 0;
  for (let i = 0; i < spans.length; i++) {
    pieces.push(xml.slice(cursor, spans[i].start));
    let serXml = xml.slice(spans[i].start, spans[i].end);
    const xLength = Math.max(1, finalScatterXLength(req, i));
    const yLength = Math.max(1, req.series?.[i]?.values.length ?? req.existing.series[i]?.values.length ?? 1);
    const xCol = chartColLetter(i * 2);
    const yCol = chartColLetter(i * 2 + 1);
    const xVal = findChartElement(serXml, 'xVal', scope);
    if (xVal) {
      serXml = serXml.slice(0, xVal.start) +
        patchFormulaRange(serXml.slice(xVal.start, xVal.end), `Sheet1!$${xCol}$2:$${xCol}$${xLength + 1}`, scope) +
        serXml.slice(xVal.end);
    }
    const yVal = findChartElement(serXml, 'yVal', scope);
    if (yVal) {
      serXml = serXml.slice(0, yVal.start) +
        patchFormulaRange(serXml.slice(yVal.start, yVal.end), `Sheet1!$${yCol}$2:$${yCol}$${yLength + 1}`, scope) +
        serXml.slice(yVal.end);
    }
    pieces.push(serXml);
    cursor = spans[i].end;
  }
  pieces.push(xml.slice(cursor));
  return pieces.join('');
}

function titleBlock(title: string, parentXml: string): string {
  const titleTag = chartQualifiedName(parentXml, 'title');
  const txTag = chartQualifiedName(parentXml, 'tx');
  const richTag = chartQualifiedName(parentXml, 'rich');
  return (
    `<${titleTag}>\n` +
    `      <${txTag}>\n` +
    `        <${richTag}>\n` +
    '          <a:bodyPr/>\n' +
    '          <a:lstStyle/>\n' +
    '          <a:p>\n' +
    '            <a:r>\n' +
    '              <a:rPr b="1" sz="1400"/>\n' +
    `              <a:t>${escapeXml(title)}</a:t>\n` +
    '            </a:r>\n' +
    '          </a:p>\n' +
    `        </${richTag}>\n` +
    `      </${txTag}>\n` +
    `    </${titleTag}>`
  );
}

function insertTitleElement(xml: string, title: string, chartPath: string, scope: ChartNamespaceScope): string {
  const chartSpan = findChartElement(xml, 'chart', scope);
  if (!chartSpan) {
    throw mergeRefusal(chartPath, 'the chart part has no <c:chart> element to carry a title', REINDEX_REMEDY);
  }
  const chartXml = xml.slice(chartSpan.start, chartSpan.end);
  const prefix = chartPrefix(chartXml);
  const markers = [
    'autoTitleDeleted',
    'pivotFmts',
    'view3D',
    'floor',
    'sideWall',
    'backWall',
    'plotArea',
    'legend',
    'plotVisOnly',
    'dispBlanksAs',
    'showDLblsOverMax',
  ];
  let at = -1;
  for (const marker of markers) {
    const tag = findChartDirectChild(chartXml, marker, scope);
    if (tag && (at < 0 || tag.start < at)) at = tag.start;
  }
  const absolute = at >= 0 ? chartSpan.start + at : chartSpan.end - `</${prefix}chart>`.length;
  return `${xml.slice(0, absolute)}${titleBlock(title, chartXml)}\n    ${xml.slice(absolute)}`;
}

function patchTitleText(xml: string, title: string, chartPath: string, scope: ChartNamespaceScope): string {
  // Chart-level title only: axis titles also use <c:title> but live inside
  // <c:plotArea>; a nested title before it is not a legal CT_Chart child.
  const chartSpan = findChartElement(xml, 'chart', scope);
  const titleSpan = chartSpan ? findDirectChartTitleSpan(xml, chartSpan, scope) : null;
  if (!titleSpan) return insertTitleElement(xml, title, chartPath, scope);
  const block = xml.slice(titleSpan.start, titleSpan.end);
  const gt = tagEnd(block, 0);
  if (gt > 0 && block[gt - 1] === '/') {
    return xml.slice(0, titleSpan.start) + titleBlock(title, block) + xml.slice(titleSpan.end);
  }
  const texts = findAllElements(block, 'a:t');
  if (texts.length === 0) {
    const tx = titleBlock(title, block).replace(/^<[^>]+>|<\/[^>]+>$/g, '');
    // A linked title (`<c:tx><c:strRef>...`) has no drawn text run; replace
    // its tx wholesale so CT_Title never carries two `c:tx` children.
    const existingTx = findChartDirectChild(block, 'tx', scope);
    if (existingTx) {
      return (
        xml.slice(0, titleSpan.start) +
        block.slice(0, existingTx.start) +
        tx +
        block.slice(existingTx.end) +
        xml.slice(titleSpan.end)
      );
    }
    const openEnd = gt + 1;
    return xml.slice(0, titleSpan.start) + block.slice(0, openEnd) + tx + block.slice(openEnd) + xml.slice(titleSpan.end);
  }
  const current = decodeXmlText(
    texts
      .map((span) => block.slice(span.start, span.end).replace(/^<a:t[^>]*>/, '').replace(/<\/a:t>$/, ''))
      .join('')
  );
  if (current === title) return xml;
  const pieces: string[] = [block.slice(0, texts[0].start)];
  for (let i = 0; i < texts.length; i++) {
    const el = block.slice(texts[i].start, texts[i].end);
    pieces.push(setElementText(el, i === 0 ? title : ''));
    pieces.push(i + 1 < texts.length ? block.slice(texts[i].end, texts[i + 1].start) : '');
  }
  pieces.push(block.slice(texts[texts.length - 1].end));
  return xml.slice(0, titleSpan.start) + pieces.join('') + xml.slice(titleSpan.end);
}

/**
 * Removes the chart-level <c:title> for an explicit empty title. Axis titles
 * also use <c:title> but live inside <c:plotArea>, so the search is bounded
 * exactly like patchTitleText. Leading whitespace is removed with the element
 * so the result matches a structurally generated titleless part.
 */
function removeTitleElement(xml: string, scope: ChartNamespaceScope): string {
  let out = xml;
  // CT_Chart allows at most one title; remove every direct title before the
  // plot area so a hostile duplicate cannot resurface after a clear.
  for (;;) {
    const chartSpan = findChartElement(out, 'chart', scope);
    if (!chartSpan) return out;
    const titleSpan = findDirectChartTitleSpan(out, chartSpan, scope);
    if (!titleSpan) return out;
    let start = titleSpan.start;
    while (start > chartSpan.start && /\s/.test(out[start - 1])) start--;
    out = out.slice(0, start) + out.slice(titleSpan.end);
  }
}

function clearAutoTitleDeleted(xml: string, scope: ChartNamespaceScope): string {
  const chartSpan = findChartElement(xml, 'chart', scope);
  if (!chartSpan) return xml;
  const chartXml = xml.slice(chartSpan.start, chartSpan.end);
  const span = findChartDirectChild(chartXml, 'autoTitleDeleted', scope);
  if (!span) return xml;
  const start = chartSpan.start + span.start;
  const end = chartSpan.start + span.end;
  return xml.slice(0, start) + setAttribute(xml.slice(start, end), 'val', '0') + xml.slice(end);
}

function insertView3DProperty(block: string, tag: string, value: string, scope: ChartNamespaceScope): string {
  const position = VIEW_3D_ORDER.indexOf(tag);
  const qualified = chartQualifiedName(block, tag);
  for (const later of VIEW_3D_ORDER.slice(position + 1)) {
    const el = findChartElement(block, later, scope);
    if (el) return `${block.slice(0, el.start)}<${qualified} val="${value}"/>\n      ${block.slice(el.start)}`;
  }
  const close = block.lastIndexOf(`</${chartQualifiedName(block, 'view3D')}>`);
  if (close < 0) return expandSelfClosedElement(block, `      <${qualified} val="${value}"/>\n    `) ?? block;
  return `${block.slice(0, close)}      <${qualified} val="${value}"/>\n    ${block.slice(close)}`;
}

function patchView3D(xml: string, view3D: View3DProperties, chartPath: string, scope: ChartNamespaceScope): string {
  const span = findChartElement(xml, 'view3D', scope);
  if (!span) {
    throw mergeRefusal(
      chartPath,
      'the 3D chart part has no <c:view3D> block to merge camera settings into',
      'omit view3D to keep the existing camera, or delete and re-insert the 3D chart, then retry.'
    );
  }
  let block = xml.slice(span.start, span.end);
  const entries: Array<[string, string]> = [];
  if (view3D.rotX !== undefined) entries.push(['rotX', String(clampNumber(view3D.rotX, -90, 90))]);
  if (view3D.hPercent !== undefined) entries.push(['hPercent', String(view3D.hPercent)]);
  if (view3D.rotY !== undefined) entries.push(['rotY', String(clampNumber(view3D.rotY, 0, 360))]);
  if (view3D.depthPercent !== undefined) entries.push(['depthPercent', String(view3D.depthPercent)]);
  if (view3D.rAngAx !== undefined) entries.push(['rAngAx', view3D.rAngAx ? '1' : '0']);
  if (view3D.perspective !== undefined) entries.push(['perspective', String(clampNumber(view3D.perspective, 0, 240))]);
  for (const [tag, value] of entries) {
    const el = findChartElement(block, tag, scope);
    block = el
      ? block.slice(0, el.start) + setAttribute(block.slice(el.start, el.end), 'val', value) + block.slice(el.end)
      : insertView3DProperty(block, tag, value, scope);
  }
  return xml.slice(0, span.start) + block + xml.slice(span.end);
}

function patchShape(xml: string, shape: Bar3DShape, chartPath: string, scope: ChartNamespaceScope): string {
  const bar = findChartElement(xml, 'bar3DChart', scope);
  if (!bar) {
    throw mergeRefusal(
      chartPath,
      'the 3D chart part has no <c:bar3DChart> to merge a bar shape into',
      'omit shape for this chart type, or use chartType column3D|bar3D, or delete and re-insert the chart.'
    );
  }
  let block = xml.slice(bar.start, bar.end);
  const shapeTag = chartQualifiedName(block, 'shape');
  const el = findChartDirectChild(block, 'shape', scope);
  if (el) {
    block = block.slice(0, el.start) + setAttribute(block.slice(el.start, el.end), 'val', shape) + block.slice(el.end);
  } else {
    const ax = findChartDirectChild(block, 'axId', scope);
    if (ax) {
      block = `${block.slice(0, ax.start)}<${shapeTag} val="${shape}"/>\n        ${block.slice(ax.start)}`;
    } else {
      const close = block.lastIndexOf(`</${chartQualifiedName(block, 'bar3DChart')}>`);
      if (close < 0) {
        throw mergeRefusal(chartPath, 'the bar3D chart block is not closed', REINDEX_REMEDY);
      }
      block = `${block.slice(0, close)}        <${shapeTag} val="${shape}"/>\n      ${block.slice(close)}`;
    }
  }
  return xml.slice(0, bar.start) + block + xml.slice(bar.end);
}

function patchHoleSize(xml: string, holeSize: number, chartPath: string, scope: ChartNamespaceScope): string {
  const target = findChartNamespaceElements(xml, ['doughnutChart'])[0];
  if (!target || target.end < 0) {
    throw mergeRefusal(
      chartPath,
      'the chart part has no <c:doughnutChart> to merge holeSize into',
      'use chartType doughnut for this chart, or omit holeSize.'
    );
  }
  const doughnut = { start: target.start, end: target.end };
  const existing = findChartElement(xml, 'holeSize', scope, doughnut.start, doughnut.end);
  if (existing) {
    return xml.slice(0, existing.start) +
      setAttribute(xml.slice(existing.start, existing.end), 'val', String(holeSize)) +
      xml.slice(existing.end);
  }
  const block = xml.slice(doughnut.start, doughnut.end);
  const ext = findChartDirectChild(block, 'extLst', scope);
  const close = `</${target.qualifiedName}>`;
  const at = ext ? doughnut.start + ext.start : xml.lastIndexOf(close);
  if (at < doughnut.start) {
    throw mergeRefusal(
      chartPath,
      'the doughnut chart block is not closed',
      REINDEX_REMEDY
    );
  }
  const holeTag = chartQualifiedName(block, 'holeSize');
  return xml.slice(0, at) + `        <${holeTag} val="${holeSize}"/>\n      ` + xml.slice(at);
}

/**
 * Applies a data-only chart update onto an existing chart part, returning the
 * patched XML. Throws E_VALIDATION (+ remedy) for any shape that cannot be
 * merged; callers must not write the part when this throws.
 */
export function mergeChartDataInPlace(chartXml: string, req: ChartMergeRequest): string {
  let xml = chartXml;
  const scope = chartNamespaceScopeFromXml(xml);
  if (req.title !== undefined) {
    if (req.title === '') {
      // An explicit empty title is a titleless part: remove the element
      // instead of leaving an empty <c:title> behind (or inserting one).
      xml = removeTitleElement(xml, scope);
    } else {
      xml = patchTitleText(xml, req.title, req.chartPath, scope);
      xml = clearAutoTitleDeleted(xml, scope);
    }
  }
  if (req.series !== undefined) {
    xml = patchSeriesList(xml, req, req.series, scope);
  }
  if (req.categories !== undefined && !sameStrings(req.existing.categories, req.categories)) {
    // Scatter series with an explicit xValues list are patched per series
    // above; categories only drives series that still track the shared list.
    const scatterPatchesAllOwnX = req.chartType === 'scatter'
      && req.series !== undefined
      && req.series.length > 0
      && req.series.every((patch) => patch.xValues !== undefined);
    xml = req.chartType === 'scatter'
      ? (scatterPatchesAllOwnX
          ? xml
          : patchScatterXValues(xml, req, req.categories, scope))
      : patchCategories(xml, req, req.categories, scope);
  }
  // Formula re-pointing assumes the workbook is regenerated in the canonical
  // layout; a data-free edit (title only) leaves both bytes untouched.
  if (req.chartType === 'scatter' && (req.categories !== undefined || req.series !== undefined)) {
    xml = patchScatterFormulas(xml, req, scope);
  }
  if (req.is3D && req.view3D !== undefined) {
    xml = patchView3D(xml, req.view3D, req.chartPath, scope);
  }
  if (req.is3D && req.shape !== undefined) {
    xml = patchShape(xml, req.shape, req.chartPath, scope);
  }
  if (!req.is3D && req.holeSize !== undefined) {
    xml = patchHoleSize(xml, req.holeSize, req.chartPath, scope);
  }
  return xml;
}
