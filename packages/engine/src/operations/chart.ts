import { DOMParser } from '@xmldom/xmldom';
import * as crypto from 'crypto';
import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node, Section } from '../ast';
import { DocxParser, detectChartKindFromXml } from '../parser';
import { insertIntoSection, lastSection, locateNode, ownerDocumentOf } from '../parser/sections';
import {
  chartElementDescendants,
  chartPointTextsByIndex,
  chartScatterXValuesBySeries,
  chartTitleElement,
  chartNamespaceScopeFromXml,
  findChartNamespaceElements,
  findPlotElementLocalNames,
  PLOT_LOCAL_NAMES_2D,
  PLOT_LOCAL_NAMES_3D,
  PLOT_LOCAL_NAMES_UNSUPPORTED,
  pairScatterPointTexts,
  scatterSeriesXList,
  scatterXOnlyValues,
} from '../parser/factory';
import type { StyleGuideConfig } from '../style-guide/types';
import { expandShorthandHex, isShorthandHex } from '../style-guide/parser';
import { resolveReproducibleMode } from '../zip-hardening';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { BaseOperation, InsertChainHost, validateVersion } from './base';
import {
  Chart3DManipulator,
  Chart3DType,
  SUPPORTED_CHART_TYPES,
  Bar3DShape,
  View3DProperties,
  Chart3DSeries,
  defaultSeriesColor,
  generateEmbeddedWorkbook
} from '../manipulators/chart3d';
import {
  Chart2DManipulator,
  Chart2DType,
  SUPPORTED_CHART_2D_TYPES,
  Chart2DSeries,
  assertScatterData
} from '../manipulators/chart2d';
import { mergeChartDataInPlace } from './chart-merge';
import { type ChartStyle, applyChartFloating, applyChartStyle, readChartStyle, validateChartStyle } from './chart-style';
import { chartSpans, findDirectChild, attrValue, chartQualifiedName } from './chart-xml-spans';

/**
 * Unified chart type system: 2D DrawingML types are the default;
 * 3D types remain as explicit opt-ins (unchanged 3D output paths).
 */
export type ChartType = Chart2DType | Chart3DType;
export const SUPPORTED_CHART_2D_AND_3D_TYPES = [...SUPPORTED_CHART_2D_TYPES, ...SUPPORTED_CHART_TYPES] as const;

export function is3DChartType(chartType: string): boolean {
  return (SUPPORTED_CHART_TYPES as readonly string[]).includes(chartType);
}

/**
 * A base type names the whole family, not a shape: column/bar on an existing
 * stacked/percent part keeps that part's grouping, and pie on an existing
 * doughnut keeps the authored hole (the pie family selector mirrors the bar
 * rule). Explicit bar variant names still select a variant.
 */
function isBaseFamilyOf(base: string, variant: string): boolean {
  return (base === 'column' && (variant === 'columnStacked' || variant === 'columnPercentStacked')) ||
    (base === 'bar' && (variant === 'barStacked' || variant === 'barPercentStacked')) ||
    (base === 'pie' && variant === 'doughnut');
}

function holeSizeValidationError(holeSize: number | undefined, chartType: string): string | undefined {
  if (holeSize === undefined) return undefined;
  if (chartType !== 'doughnut') return 'holeSize is only supported for doughnut charts.';
  if (!Number.isInteger(holeSize) || holeSize < 1 || holeSize > 90) return 'holeSize must be an integer from 1 to 90.';
  return undefined;
}

function isChartReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

function chartHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

function chartDeterministicId(seed: string, ast: DocumentAST): string {
  const base = `chart_${chartHash8(seed)}`;
  const live = new Set<string>();
  try {
    for (const s of ast.sections ?? []) {
      for (const n of s.nodes ?? []) {
        if (typeof n.id === 'string' && n.id) live.add(n.id);
      }
    }
  } catch { /* best-effort */ }
  let candidate = base;
  let n = 2;
  while (live.has(candidate)) candidate = `${base}_${n++}`;
  return candidate;
}

function chartDeterministicChecksum(seed: string): string {
  if (!isChartReproducible()) return `chart_chk_${Date.now()}`;
  try {
    return `chart_chk_${chartHash8(seed)}`;
  } catch {
    return 'chart_chk_00000000';
  }
}

/**
 * Guide chart palette as clean 6-char uppercase hex, or undefined when the
 * guide carries no usable palette. Shared by insert + update so both fill
 * only uncolored series in order (explicit colors keep their indices).
 */
function cleanGuidePalette(guide?: StyleGuideConfig): string[] | undefined {
  const raw = guide?.palette?.chartPalette;
  if (!Array.isArray(raw) || raw.length === 0) return undefined;
  const clean = raw
    .map((c) => String(c ?? '').replace(/^#/, '').toUpperCase())
    .filter((c) => /^[0-9A-F]{6}$/.test(c));
  return clean.length > 0 ? clean : undefined;
}

/**
 * Mixes a 6-char hex color toward white (positive amount) or black
 * (negative amount); magnitude is clamped to [0, 1]. Used to derive extra
 * series colors once a guide palette is exhausted, so long series lists
 * never wrap around to an already-used palette entry.
 */
function shadeHex(hex: string, amount: number): string {
  const num = parseInt(hex.replace(/^#/, '').toUpperCase(), 16);
  const target = amount < 0 ? 0 : 255;
  const p = Math.min(1, Math.abs(amount));
  const mix = (c: number): number =>
    Math.max(0, Math.min(255, Math.round(c + (target - c) * p)));
  const parts = [mix((num >> 16) & 0xff), mix((num >> 8) & 0xff), mix(num & 0xff)];
  return parts.map((c) => c.toString(16).toUpperCase().padStart(2, '0')).join('');
}

/**
 * Palette entry for a series index. Inside the palette this is the entry
 * itself; past the end it is a generated lighter/darker shade of the
 * wrapped entry (alternating direction, growing strength), skipping shades
 * that collide with the base, any palette entry, or any previously
 * generated overflow color in `taken`. Callers threading one set across
 * indices keep every overflow output distinct; the fallback sweep below is
 * collision-checked the same way (the trailing return is unreachable in
 * practice — 16M colors — but still valid 6-hex).
 */
function guidePaletteColor(palette: string[], index: number, taken?: Set<string>): string {
  if (index < palette.length) return palette[index];
  const seen = taken ?? new Set<string>();
  for (const c of palette) seen.add(c.toUpperCase());
  const base = palette[index % palette.length].toUpperCase();
  const depth = Math.floor(index / palette.length);
  const step = 0.25 * Math.ceil(depth / 2);
  const firstDir = depth % 2 === 1 ? -1 : 1;
  for (const dir of [firstDir, -firstDir]) {
    let amount = dir * step;
    while (Math.abs(amount) < 1) {
      const out = shadeHex(base, amount);
      if (out !== base && !seen.has(out)) {
        seen.add(out);
        return out;
      }
      amount += dir * 0.15;
    }
  }
  for (let p = 0.05; p < 1; p += 0.05) {
    for (const dir of [firstDir, -firstDir]) {
      const out = shadeHex(base, dir * p);
      if (out !== base && !seen.has(out)) {
        seen.add(out);
        return out;
      }
    }
  }
  const last = shadeHex(base, firstDir * 0.97);
  seen.add(last);
  return last;
}

/**
 * Fills uncolored series from the guide palette in order; explicitly
 * colored series pass through untouched (their indices never shift the
 * rest). Without a usable palette the input ref returns as-is, so
 * guide-absent output stays byte-identical to the legacy path.
 */
function withGuideSeriesColors<T extends { color?: unknown }>(
  series: T[] | undefined,
  palette: string[] | undefined,
): T[] | undefined {
  if (!series || !palette || palette.length === 0) return series;
  const taken = new Set<string>();
  return series.map((s, i) => {
    const raw = typeof s?.color === 'string' ? s.color.replace(/^#/, '') : '';
    const existing = isShorthandHex(raw) ? expandShorthandHex(raw) : raw;
    if (/^[0-9A-Fa-f]{6}$/.test(existing)) return s;
    return { ...s, color: guidePaletteColor(palette, i, taken) };
  });
}

export interface ChartOperation extends BaseOperation {
  action: 'chart';
  operation?: 'insert' | 'update' | 'updateView3D';
  chartType?: ChartType;
  /** Doughnut hole diameter percentage (integer 1..90; default 50). */
  holeSize?: number;
  title?: string;
  view3D?: View3DProperties;
  shape?: Bar3DShape;
  categories?: string[];
  series?: Chart3DSeries[] | Chart2DSeries[];
  style?: ChartStyle;
  anchor?: string;
  id?: string;
  target?: string;
  widthEmu?: number;
  heightEmu?: number;
  params?: Record<string, unknown>;
}

export function resolveChartOp(op: ChartOperation): ChartOperation {
  const raw = ((op as unknown as { params?: Record<string, unknown> }).params || {}) as Record<string, unknown>;
  return {
    ...raw,
    ...op,
    action: 'chart',
    operation: op.operation ?? (raw.operation as ChartOperation['operation']),
    target: op.target ?? (raw.target as string | undefined),
    chartType: op.chartType ?? (raw.chartType as ChartType | undefined),
    holeSize: op.holeSize ?? (raw.holeSize as number | undefined),
    title: op.title ?? (raw.title as string | undefined),
    categories: op.categories ?? (raw.categories as string[] | undefined),
    series: op.series ?? (raw.series as ChartOperation['series']),
    style: op.style ?? (raw.style as ChartStyle | undefined),
    anchor: op.anchor ?? (raw.anchor as string | undefined),
    id: op.id ?? (raw.id as string | undefined),
    view3D: op.view3D ?? (raw.view3D as View3DProperties | undefined),
    shape: op.shape ?? (raw.shape as Bar3DShape | undefined),
    widthEmu: op.widthEmu ?? (raw.widthEmu as number | undefined),
    heightEmu: op.heightEmu ?? (raw.heightEmu as number | undefined),
    expectedVersion: op.expectedVersion ?? (raw.expectedVersion as number | undefined),
  };
}

export function executeChart(op: ChartOperation, ast: DocumentAST, parser?: DocxParser, guide?: StyleGuideConfig): ChartUpdateResult | void {
  const resolved = resolveChartOp(op);
  const manipulator3D = new Chart3DManipulator(ast, parser);

  if (resolved.operation === 'update') {
    return executeChartUpdate(resolved, ast, parser, guide);
  }

  if (resolved.operation === 'updateView3D') {
    if (resolved.style !== undefined) {
      throw Object.assign(
        new Error('E_VALIDATION: chart style cannot be applied through updateView3D.'),
        {
          code: 'E_VALIDATION',
          retryable: false,
          remedy: "Use operation:'update' with style, or call updateView3D without style and then a separate style-only update.",
        }
      );
    }
    if (parser) {
      const zipAdapter = parser.getZipAdapter();
      const chartPath = 'word/charts/chart1.xml';
      if (zipAdapter.hasFile(chartPath)) {
        const existingXml = zipAdapter.readText(chartPath);
        assertNoXxe(existingXml, chartPath);
        const updatedXml = manipulator3D.updateView3D(existingXml, resolved.view3D || {});
        zipAdapter.writeText(chartPath, updatedXml);
      }
    }
    return;
  }

  const rawChartType = resolved.chartType || 'column';
  const supported2DTypes = SUPPORTED_CHART_2D_TYPES as readonly string[];
  const supported3DTypes = SUPPORTED_CHART_TYPES as readonly string[];
  if (!supported2DTypes.includes(rawChartType) && !supported3DTypes.includes(rawChartType)) {
    throw Object.assign(
      new Error(`Unsupported chart type "${rawChartType}". Supported types: ${SUPPORTED_CHART_2D_AND_3D_TYPES.join(', ')}`),
      {
        code: 'E_VALIDATION',
        retryable: false,
        remedy: `Use one of the supported chart types: ${SUPPORTED_CHART_2D_AND_3D_TYPES.join(', ')}`,
      }
    );
  }

  const use3D = is3DChartType(rawChartType);
  assertComboSeriesFields(resolved.series as Array<{ chartType?: unknown; axis?: unknown }> | undefined, rawChartType, 'new chart part');
  const holeErr = holeSizeValidationError(resolved.holeSize, rawChartType);
  if (holeErr) {
    throw Object.assign(new Error(holeErr), {
      code: 'E_VALIDATION',
      retryable: false,
      remedy: 'Use an integer holeSize from 1 to 90 for doughnut charts, or omit it for the default 50.',
    });
  }
  if (rawChartType === 'scatter') {
    assertScatterData(resolved.categories, resolved.series as Chart2DSeries[] | undefined);
  }
  const defaultTitle = use3D ? '3D Chart Analysis' : 'Chart Analysis';
  const fallbackText = use3D ? 'Q1-Q4 Financial 3D Analysis' : 'Q1-Q4 Financial Analysis';
  const fallbackTitle = use3D ? 'Time Saved by SDLC Stage (Hours)' : 'Time Saved by SDLC Stage (Hours)';

  // Same-anchor 'after' inserts preserve FIFO order via the per-anchor
  // chain shared with insert/equation/callout (same _insertAfterChain /
  // _lastInsertAfter key/shape): re-validated against live DOM/AST at use,
  // dead links dropped so the next insert lands after the latest live link.
  // The tail id is forwarded as the DOM anchor so the drawing lands after
  // the tail in both AST and DOM.
  let chainHost: Node | undefined;
  let liveChain: Node[] = [];
  let effectiveAnchor: Node | undefined;
  if (resolved.anchor) {
    for (const section of ast.sections) {
      const found = section.nodes.find((n) => n.id === resolved.anchor);
      if (found) {
        chainHost = found;
        break;
      }
    }
    if (chainHost) {
      const chainCaps = chainHost as unknown as InsertChainHost;
      const rawChain: Node[] = Array.isArray(chainCaps._insertAfterChain)
        ? chainCaps._insertAfterChain as Node[]
        : [];
      const legacyTail = chainCaps._lastInsertAfter;
      if (rawChain.length === 0 && legacyTail) {
        rawChain.push(legacyTail);
      }
      liveChain = rawChain.filter((n) => {
        if (!n || !n.xml || !(n.xml as Element).parentNode) return false;
        return ast.sections.some((s: Section) => s.nodes.some((x: Node) => x.id === n.id));
      });
      chainCaps._insertAfterChain = liveChain;
      if (legacyTail && !liveChain.some((n) => n.id === legacyTail.id)) {
        delete chainCaps._lastInsertAfter;
      }
      effectiveAnchor = liveChain.length > 0 ? liveChain[liveChain.length - 1] : chainHost;
    }
  }
  const domAnchorId = effectiveAnchor ? effectiveAnchor.id : resolved.anchor;

  // Execute chart insertion via the routed manipulator (2D default, 3D opt-in).
  // With a guide active, uncolored series pull from the guide chart palette
  // in order (explicit colors keep their indices); without one the series
  // pass through untouched (legacy defaults apply downstream).
  const guidedSeries = withGuideSeriesColors(
    resolved.series as Array<{ color?: unknown }> | undefined,
    cleanGuidePalette(guide),
  ) as ChartOperation['series'];
  // Scatter resolves one x list per series up front so the chart part, AST
  // metadata, TOON/HTML and the embedded workbook all carry the same x data,
  // including when only some series supply xValues.
  let preparedSeries = guidedSeries as Chart2DSeries[] | undefined;
  let scatterXValues: number[][] | undefined;
  if (rawChartType === 'scatter') {
    scatterXValues = assertScatterData(resolved.categories, preparedSeries);
    preparedSeries = preparedSeries!.map((s, i) => ({ ...s, xValues: scatterXValues![i] }));
  }
  if (rawChartType === 'combo' && preparedSeries !== undefined) {
    preparedSeries = preparedSeries.map((s) => ({
      ...s,
      chartType: s.chartType ?? 'column',
      axis: s.axis ?? 'primary',
    }));
  }
  const styleContext = {
    chartPath: 'new chart part',
    chartType: rawChartType,
    seriesCount: preparedSeries?.length ?? 0,
    is3D: use3D,
  };
  validateChartStyle(resolved.style, styleContext);
  const transformChartXml = resolved.style !== undefined
    ? (xml: string): string => applyChartStyle(xml, resolved.style, styleContext)
    : undefined;

  let drawingNodeId: string | undefined;
  if (use3D) {
    const chartRes = manipulator3D.insert3DChart({
      chartType: rawChartType as Chart3DType,
      title: resolved.title || defaultTitle,
      view3D: resolved.view3D,
      shape: resolved.shape,
      categories: resolved.categories,
      series: preparedSeries as Chart3DSeries[] | undefined,
      anchorNodeId: domAnchorId,
      id: resolved.id,
      widthEmu: resolved.widthEmu,
      heightEmu: resolved.heightEmu,
      ...(transformChartXml ? { transformChartXml } : {}),
      ...(resolved.style?.floating ? { floating: resolved.style.floating } : {})
    });
    drawingNodeId = chartRes.drawingNodeId;
  } else {
    const manipulator2D = new Chart2DManipulator(ast, parser);
    const chartRes = manipulator2D.insert2DChart({
      chartType: rawChartType as Chart2DType,
      title: resolved.title || defaultTitle,
      categories: resolved.categories,
      series: preparedSeries,
      holeSize: resolved.holeSize,
      anchorNodeId: domAnchorId,
      id: resolved.id,
      widthEmu: resolved.widthEmu,
      heightEmu: resolved.heightEmu,
      ...(transformChartXml ? { transformChartXml } : {}),
      ...(resolved.style?.floating ? { floating: resolved.style.floating } : {})
    });
    drawingNodeId = chartRes.drawingNodeId;
  }

  // Sync AST node list so UI & engine immediately reflect the new chart
  if (ast.sections.length > 0) {
    let chartXmlNode: Element | undefined;
    if (parser && drawingNodeId) {
      chartXmlNode = parser.getNode(drawingNodeId);
    }
    if (!chartXmlNode) {
      const doc = ownerDocumentOf(ast.sections) || new DOMParser().parseFromString('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', 'text/xml');
      chartXmlNode = createWordElement(doc, 'w:p');
    }

    const chartSeed = `${rawChartType}:${resolved.title ?? fallbackTitle}:${JSON.stringify(resolved.categories ?? [])}:${JSON.stringify(resolved.series ?? [])}:${domAnchorId ?? ''}`;
    const insertedSeries = preparedSeries || [{ name: 'Hours Saved', values: [24, 32, 65, 48, 18], color: '2563EB' }];
    // For scatter every series carries its resolved x list; categories mirror
    // the shared input (or the first series' x when no categories were given).
    const insertedCategories = scatterXValues !== undefined
      ? (scatterXValues[0] ?? []).map((value) => String(value))
      : resolved.categories || ['Requirements', 'Design', 'Implementation', 'Testing', 'Deployment'];
    const newNode: Node = {
      id: resolved.id || drawingNodeId || (isChartReproducible() ? chartDeterministicId(chartSeed, ast) : `chart_${Date.now()}`),
      type: use3D ? 'chart3d' : 'chart',
      xml: chartXmlNode,
      metadata: {
        text: resolved.title || fallbackText,
        title: resolved.title || fallbackTitle,
        categories: insertedCategories,
        series: insertedSeries,
        chartType: resolved.chartType || 'column',
        chartData: {
          title: resolved.title || 'Time Saved by SDLC Stage (Hours)',
          categories: insertedCategories,
          series: insertedSeries,
          chartType: resolved.chartType || 'column',
          ...(rawChartType === 'doughnut' ? { holeSize: resolved.holeSize ?? 50 } : {}),
          ...(resolved.style !== undefined ? { style: resolved.style } : {})
        },
        hasCitation: false,
        hasComment: false,
        hasBookmark: false,
        hasTrackedChanges: false,
        hasEquation: false,
        hasField: false,
        hasImage: false,
        hasHyperlink: false
      },
      version: 1,
      checksum: chartDeterministicChecksum(chartSeed)
    };

    // After the anchor (or the last block chained behind it), wherever its section is;
    // with no anchor, at the end of the document.
    const placeAfter =
      (chainHost && effectiveAnchor ? locateNode(ast.sections, effectiveAnchor.id) : undefined) ??
      (resolved.anchor ? locateNode(ast.sections, resolved.anchor) : undefined);
    if (placeAfter) {
      insertIntoSection(ast.sections, placeAfter.section, placeAfter.index + 1, [newNode]);
    } else {
      lastSection(ast.sections)!.nodes.push(newNode);
    }
    if (resolved.anchor && chainHost && effectiveAnchor) {
      liveChain.push(newNode);
      (chainHost as unknown as InsertChainHost)._insertAfterChain = liveChain;
      (chainHost as unknown as InsertChainHost)._lastInsertAfter = newNode;
    }
  }
}

/** Metadata follows the part, keeping previously read fields the part still carries. */
function mergeUpdatedChartStyle(previous: unknown, requested: ChartStyle, updatedXml: string): ChartStyle | undefined {
  const merged: ChartStyle = { ...(readChartStyle(updatedXml) ?? {}) };
  const previousStyle = previous !== null && typeof previous === 'object' ? previous as ChartStyle : undefined;
  if (requested.trendline === undefined && merged.trendline === undefined && previousStyle?.trendline !== undefined) {
    merged.trendline = previousStyle.trendline;
  }
  if (requested.floating !== undefined) {
    merged.floating = requested.floating;
  } else if (merged.floating === undefined && previousStyle?.floating !== undefined) {
    merged.floating = previousStyle.floating;
  }
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function chartUpdateError(message: string, remedy: string): Error {
  return Object.assign(new Error(message), {
    code: 'E_VALIDATION',
    retryable: false,
    remedy,
  });
}

/** Per-series plot/axis selectors are meaningful only inside a combo chart. */
function assertComboSeriesFields(
  series: Array<{ chartType?: unknown; axis?: unknown }> | undefined,
  chartType: string,
  chartPath: string
): void {
  if (chartType === 'combo' || series === undefined) return;
  const index = series.findIndex((entry) => entry.chartType !== undefined || entry.axis !== undefined);
  if (index < 0) return;
  throw chartUpdateError(
    `E_VALIDATION: series[${index}] carries chartType/axis but chart part ${chartPath} is "${chartType}", not "combo".`,
    'Remedy: omit series[].chartType and series[].axis for single-plot charts, or set chartType:"combo" with at least one column and one line series.'
  );
}

interface ExistingChartSeries {
  name: string;
  values: number[];
  color?: string;
  xValues?: number[];
  unpairedX?: number;
  unpairedY?: number;
  chartType?: 'column' | 'line';
  axis?: 'primary' | 'secondary';
  idx?: number;
}

interface ExistingChartData {
  title: string;
  categories: string[];
  series: ExistingChartSeries[];
  externalDataRId?: string;
  view3D: View3DProperties;
  shape?: Bar3DShape;
  holeSize?: number;
  detected: { is3D: boolean; chartType: string };
}

/** Scatter merges rewrite point lists positionally, so sparse idx lists refuse. */
function assertDenseScatterPoints(chartXml: string, chartPath: string): void {
  const doc = new DOMParser().parseFromString(chartXml, 'text/xml');
  const root = doc?.documentElement as Element | undefined;
  if (!root) return;
  for (const ser of chartElementDescendants(root, 'ser')) {
    for (const tag of ['xVal', 'val', 'yVal']) {
      const container = chartElementDescendants(ser, tag)[0];
      if (!container) continue;
      const pts = chartElementDescendants(container, 'pt');
      for (let i = 0; i < pts.length; i++) {
        if (Number(pts[i].getAttribute('idx')) !== i) {
          throw chartUpdateError(
            `E_VALIDATION: scatter chart part ${chartPath} carries non-dense point indexes; a positional merge would move data onto the wrong points.`,
            'Remedy: re-read the chart via inspect_document; if the part was authored elsewhere, delete it and re-insert it with action:\'chart\', operation:\'insert\', chartType:\'scatter\', then apply the data update.'
          );
        }
      }
    }
  }
}

function parseExistingChartData(chartXml: string): ExistingChartData {
  const detected = detectChartKindFromXml(chartXml);
  const doc = new DOMParser().parseFromString(chartXml, 'text/xml');
  const root = doc?.documentElement as Element | undefined;
  const descendants = (localName: string): Element[] => (root ? chartElementDescendants(root, localName) : []);
  const pickCache = (parent: Element, names: string[]): Element | undefined => {
    for (const name of names) {
      const found = chartElementDescendants(parent, name)[0];
      if (found) return found;
    }
    return undefined;
  };
  const pointValues = (container: Element): string[] => chartElementDescendants(container, 'pt')
    .map((pt) => chartElementDescendants(pt, 'v')[0]?.textContent ?? '');
  const titleEl = root ? chartTitleElement(root) : undefined;
  const titleText = titleEl ? titleEl.getElementsByTagName('a:t')[0]?.textContent : undefined;
  // Never substitute a title the part does not carry; merge keeps an empty
  // title titleless and only an explicit title update writes one.
  const title = titleText ?? '';
  const isScatterPart = detected.chartType === 'scatter';
  const scatterXText = isScatterPart && root ? chartScatterXValuesBySeries(root) : [];
  const catEl = descendants('cat')[0];
  const catContainer = catEl ? pickCache(catEl, ['strCache', 'strLit', 'numCache', 'numLit']) : undefined;
  let categories = isScatterPart
    ? (scatterSeriesXList(scatterXText, 0) ?? (catContainer ? pointValues(catContainer) : []))
    : catContainer ? pointValues(catContainer) : [];
  const series: ExistingChartSeries[] = [];
  const serEls = descendants('ser');
  const combo = detected.chartType === 'combo';
  if (combo) {
    serEls.sort((a, b) => {
      const ai = Number(chartElementDescendants(a, 'idx')[0]?.getAttribute('val') ?? Number.NaN);
      const bi = Number(chartElementDescendants(b, 'idx')[0]?.getAttribute('val') ?? Number.NaN);
      return ai - bi;
    });
  }
  for (let sIdx = 0; sIdx < serEls.length; sIdx++) {
    const serEl = serEls[sIdx];
    const txEl = chartElementDescendants(serEl, 'tx')[0];
    const name = txEl ? chartElementDescendants(txEl, 'v')[0]?.textContent : undefined;
    const valEl = chartElementDescendants(serEl, 'val')[0];
    const yValEl = isScatterPart && !valEl ? chartElementDescendants(serEl, 'yVal')[0] : undefined;
    const valContainer = valEl
      ? pickCache(valEl, ['numCache', 'numLit'])
      : yValEl
        ? pickCache(yValEl, ['numCache', 'numLit'])
        : undefined;
    const xOverride = isScatterPart ? scatterSeriesXList(scatterXText, sIdx) : undefined;
    let values = valContainer ? pointValues(valContainer).map((v) => parseFloat(v) || 0) : [];
    let unpairedX = 0;
    let unpairedY = 0;
    let xValues = xOverride !== undefined ? xOverride.map((v) => parseFloat(v) || 0) : undefined;
    if (isScatterPart && xOverride !== undefined && valContainer) {
      const paired = pairScatterPointTexts(xOverride, chartPointTextsByIndex(valContainer));
      xValues = paired.xValues.length > 0 ? paired.xValues : undefined;
      values = paired.values;
      unpairedX = paired.unpairedX;
      unpairedY = paired.unpairedY;
    } else if (isScatterPart && xOverride !== undefined) {
      // A series whose y container was deleted still carries x; expose it and
      // count every non-empty point as unpaired so readers do not hide it.
      const xOnly = scatterXOnlyValues(xOverride);
      xValues = xOnly.xValues.length > 0 ? xOnly.xValues : undefined;
      unpairedX = xOnly.unpairedX;
    }
    const colorEl = serEl.getElementsByTagName('a:srgbClr')[0];
    const color = colorEl?.getAttribute('val') || undefined;
    let chartType: 'column' | 'line' | undefined;
    let axis: 'primary' | 'secondary' | undefined;
    let globalIdx: number | undefined;
    if (combo) {
      const idxText = chartElementDescendants(serEl, 'idx')[0]?.getAttribute('val');
      const parsedIdx = idxText === undefined || idxText === null ? Number.NaN : Number(idxText);
      globalIdx = Number.isInteger(parsedIdx) && parsedIdx >= 0 ? parsedIdx : undefined;
      let plot: Element | null = serEl.parentNode as Element | null;
      while (plot && plot.localName !== 'barChart' && plot.localName !== 'lineChart') plot = plot.parentNode as Element | null;
      if (plot) {
        chartType = plot.localName === 'lineChart' ? 'line' : 'column';
        const axIds = chartElementDescendants(plot, 'axId').map((el) => el.getAttribute('val'));
        axis = axIds.includes('2001') || axIds.includes('2002') ? 'secondary' : 'primary';
      }
    }
    series.push({
      name: name || 'Series',
      values,
      ...(xValues ? { xValues } : {}),
      ...(unpairedX > 0 ? { unpairedX } : {}),
      ...(unpairedY > 0 ? { unpairedY } : {}),
      ...(color ? { color: color.toUpperCase() } : {}),
      ...(chartType ? { chartType } : {}),
      ...(axis ? { axis } : {}),
      ...(globalIdx !== undefined ? { idx: globalIdx } : {}),
    });
  }
  // Scatter categories mirror the first series' paired x list, so a sparse
  // c:xVal list cannot leave holes that disagree with xValues.
  if (isScatterPart && series.length > 0 && series[0].xValues !== undefined) {
    categories = series[0].xValues.map((value) => String(value));
  }
  const extEl = descendants('externalData')[0];
  const externalDataRId = extEl?.getAttribute('r:id') || undefined;
  const view3D: View3DProperties = {};
  const viewBlock = descendants('view3D')[0];
  if (viewBlock) {
    const valFor = (tag: string): string | undefined => chartElementDescendants(viewBlock, tag)[0]?.getAttribute('val') || undefined;
    const rotX = valFor('rotX');
    const rotY = valFor('rotY');
    const perspective = valFor('perspective');
    const rAngAx = valFor('rAngAx');
    const hPercent = valFor('hPercent');
    const depthPercent = valFor('depthPercent');
    if (rotX !== undefined) view3D.rotX = parseFloat(rotX);
    if (rotY !== undefined) view3D.rotY = parseFloat(rotY);
    if (perspective !== undefined) view3D.perspective = parseFloat(perspective);
    if (rAngAx !== undefined) view3D.rAngAx = rAngAx !== '0';
    if (hPercent !== undefined) view3D.hPercent = parseFloat(hPercent);
    if (depthPercent !== undefined) view3D.depthPercent = parseFloat(depthPercent);
  }
  const shape = descendants('shape')[0]?.getAttribute('val') || undefined;
  const hole = descendants('holeSize')[0]?.getAttribute('val');
  const holeSize = hole !== undefined && hole !== null && hole !== '' ? parseInt(hole, 10) : undefined;
  return {
    title,
    categories,
    series,
    ...(externalDataRId ? { externalDataRId } : {}),
    view3D,
    ...(shape ? { shape: shape as Bar3DShape } : {}),
    ...(holeSize !== undefined && Number.isFinite(holeSize) ? { holeSize } : {}),
    detected,
  };
}

/**
 * Resolves stored scatter x for a series whose update omits xValues. The
 * stored list is matched by series name, never by array position.
 */
function resolveStoredScatterX(
  existingSeries: ExistingChartSeries[],
  name: string,
  chartPath: string
): number[] | undefined {
  const matches = existingSeries.filter((series) => series.name === name);
  if (matches.length > 1) {
    throw chartUpdateError(
      `E_VALIDATION: scatter chart part ${chartPath} has ${matches.length} series named "${name}", so stored x values cannot be matched by name.`,
      `Remedy: pass xValues for series "${name}" explicitly (one finite numeric x per y value) instead of relying on the stored list.`
    );
  }
  if (matches.length === 1) return matches[0].xValues;
  // A genuinely new series may inherit the shared x list only when every
  // stored series already tracks that same list; a divergent chart has no
  // safe list to attach to a name that is not stored.
  const sharedX = existingSeries[0]?.xValues;
  const shared = existingSeries.every((series) => {
    if (series.xValues === undefined || sharedX === undefined) return series.xValues === sharedX;
    return series.xValues.length === sharedX.length
      && series.xValues.every((value, i) => value === sharedX[i]);
  });
  if (shared) return sharedX;
  throw chartUpdateError(
    `E_VALIDATION: scatter update adds series "${name}" without xValues to chart part ${chartPath}, whose stored series carry different x lists; the new series has no stored x list to match.`,
    `Remedy: pass xValues for the new series "${name}" explicitly (one finite numeric x per y value), or omit it from this update.`
  );
}


/** Exact supported combo signature: one bar plot + one line plot on known axes. */
function assertSupportedComboXml(chartXml: string, chartPath: string): void {
  const scope = chartNamespaceScopeFromXml(chartXml);
  const directVal = (block: string, name: string): string | undefined => {
    const span = findDirectChild(block, chartQualifiedName(block, name));
    if (!span) return undefined;
    return attrValue(block.slice(span.start, span.end), 'val');
  };
  const axes = (kind: string): Array<{ id: string; hidden: boolean }> => chartSpans(chartXml, kind, scope).map((span) => {
    const block = chartXml.slice(span.start, span.end);
    return { id: directVal(block, 'axId') ?? '', hidden: directVal(block, 'delete') === '1' };
  });
  const catAxes = axes('catAx');
  const valAxes = axes('valAx');
  const plotIds = (plot: { start: number; end: number }): string[] => {
    const block = chartXml.slice(plot.start, plot.end);
    return [...block.matchAll(/<(?:[A-Za-z_][\w.-]*:)?axId\s+val="(\d+)"/g)].map((match) => match[1]);
  };
  const barPlots = findChartNamespaceElements(chartXml, ['barChart']);
  const linePlots = findChartNamespaceElements(chartXml, ['lineChart']);
  const ids = (items: Array<{ id: string }>): string[] => items.map((item) => item.id).sort();
  const primary = ids(catAxes).includes('1001') && ids(valAxes).includes('1002');
  const secondary = ids(catAxes).includes('2001') && ids(valAxes).includes('2002');
  const lineIdSets = linePlots.map((plot) => plotIds(plot).join(','));
  const lineShapeOk = linePlots.length >= 1
    && linePlots.length <= 2
    && lineIdSets.every((set) => set === '1001,1002' || set === '2001,2002')
    && new Set(lineIdSets).size === lineIdSets.length;
  const plotShapeOk = barPlots.length === 1
    && plotIds(barPlots[0]).join(',') === '1001,1002'
    && lineShapeOk
    && lineIdSets.includes('2001,2002') === secondary;
  const secondaryCatHidden = !secondary || catAxes.some((axis) => axis.id === '2001' && axis.hidden);
  const allowedCats = secondary ? ['1001', '2001'] : ['1001'];
  const allowedVals = secondary ? ['1002', '2002'] : ['1002'];
  const shapeOk = catAxes.length === allowedCats.length
    && valAxes.length === allowedVals.length
    && ids(catAxes).join(',') === allowedCats.join(',')
    && ids(valAxes).join(',') === allowedVals.join(',')
    && primary
    && secondaryCatHidden
    && plotShapeOk;
  if (!shapeOk) {
    throw chartUpdateError(
      `E_VALIDATION: chart update failed: unsupported combo axis signature in ${chartPath}.`,
      'Remedy: update supports combo parts with one barChart + one lineChart on axes 1001/1002 (plus optional hidden 2001/2002 secondary axes). Delete and re-insert the chart with action:\'chart\', chartType:\'combo\'.'
    );
  }
  const sers = chartSpans(chartXml, 'ser', scope);
  if (sers.length === 0) {
    throw chartUpdateError(
      `E_VALIDATION: chart update failed: combo chart part has no series in ${chartPath}.`,
      'Remedy: delete and re-insert the combo chart with column and line series.'
    );
  }
  const seen = new Set<number>();
  for (const span of sers) {
    const idxText = directVal(chartXml.slice(span.start, span.end), 'idx');
    const idx = idxText === undefined ? NaN : Number(idxText);
    if (!Number.isInteger(idx) || idx < 0 || idx >= sers.length || seen.has(idx)) {
      throw chartUpdateError(
        `E_VALIDATION: chart update failed: combo chart part has non-dense series indexes in ${chartPath}.`,
        'Remedy: delete and re-insert the combo chart so series indexes are dense 0..n-1, then retry the update.'
      );
    }
    seen.add(idx);
  }
}

function assertSupportedChartXml(chartXml: string, chartPath: string): void {
  const found = findPlotElementLocalNames(chartXml);
  const supported = found.filter((tag) =>
    (PLOT_LOCAL_NAMES_2D as readonly string[]).includes(tag) ||
    (PLOT_LOCAL_NAMES_3D as readonly string[]).includes(tag));
  const unsupported = found.filter((tag) => (PLOT_LOCAL_NAMES_UNSUPPORTED as readonly string[]).includes(tag));
  if (found.length > 1) {
    const combo = found.length >= 2 && found.length <= 3
      && found.includes('barChart')
      && found.includes('lineChart')
      && found.every((tag) => tag === 'barChart' || tag === 'lineChart');
    if (combo) {
      assertSupportedComboXml(chartXml, chartPath);
      return;
    }
    throw chartUpdateError(
      'E_VALIDATION: chart update failed: multi-plot chart part (' + found.join(', ') + ') in ' + chartPath + '.',
      'Remedy: update supports single-plot charts and the supported barChart+lineChart combo only. Delete the combined chart and re-insert one chart with a supported type, then retry.'
    );
  }
  if (unsupported.length > 0 || supported.length === 0) {
    const detail = found.length > 0 ? ' (' + found.join(', ') + ')' : ' (no supported plot element)';
    throw chartUpdateError(
      'E_VALIDATION: chart update failed: unsupported chart shape' + detail + ' in ' + chartPath + '.',
      'Remedy: re-read via inspect_document; update supports ' + SUPPORTED_CHART_2D_AND_3D_TYPES.join(' | ') + ' only. Delete and re-insert the chart with a supported type, then retry.'
    );
  }
  const plotTag = supported[0];
  const plot = findChartNamespaceElements(chartXml, [plotTag])[0];
  if (plot?.selfClosing) {
    throw chartUpdateError(
      'E_VALIDATION: chart update failed: empty <' + plot.qualifiedName + '/> plot element in ' + chartPath + '.',
      'Remedy: the plot element has no series. Delete and re-insert the chart with a supported type, then retry.'
    );
  }
  if (!/<(?:[A-Za-z_][\w.-]*:)?chart(?=[\s/>])/.test(chartXml)) {
    throw chartUpdateError(
      'E_VALIDATION: chart update failed: missing <c:chart> in ' + chartPath + '.',
      'Remedy: the chart part is corrupt. Delete and re-insert the chart with a supported type, then retry.'
    );
  }
}

function listChartPartNumbers(zipAdapter: { hasFile(path: string): boolean }): number[] {
  const nums: number[] = [];
  for (let n = 1; n <= 999; n++) {
    if (zipAdapter.hasFile('word/charts/chart' + n + '.xml')) nums.push(n);
  }
  return nums;
}

function listChartNodeIds(ast: DocumentAST): string[] {
  const ids: string[] = [];
  for (const section of ast.sections) {
    for (const node of section.nodes) {
      if (node.type === 'chart' || node.type === 'chart3d') ids.push(node.id);
    }
  }
  return ids;
}

function chartNumFromTarget(target: string): number | undefined {
  const m = /chart(\d+)\.xml$/i.exec(target);
  return m ? parseInt(m[1], 10) : undefined;
}

function resolveChartRIdTarget(parser: DocxParser, rId: string, ast: DocumentAST): string | undefined {
  const fromAst = (ast.relationships as Record<string, { target?: string }>)[rId]?.target;
  if (typeof fromAst === 'string' && fromAst.length > 0) return fromAst;
  const adapter = parser.getZipAdapter();
  const docRels = adapter.readTextOrNull('word/_rels/document.xml.rels');
  if (docRels) {
    assertNoXxe(docRels, 'word/_rels/document.xml.rels');
    try {
      const doc = new DOMParser().parseFromString(docRels, 'text/xml');
      const rels = Array.from(doc.getElementsByTagName('Relationship'));
      for (const rel of rels) {
        if ((rel as Element).getAttribute('Id') === rId) {
          const t = (rel as Element).getAttribute('Target') || undefined;
          if (t) return t;
        }
      }
    } catch { /* fall through to header/footer scan */ }
  }
  for (let n = 1; n <= 9; n++) {
    for (const part of ['word/_rels/header' + n + '.xml.rels', 'word/_rels/footer' + n + '.xml.rels']) {
      let xml: string | null = null;
      try {
        xml = adapter.readTextOrNull(part);
      } catch { xml = null; }
      if (!xml) continue;
      assertNoXxe(xml, part);
      try {
        const doc = new DOMParser().parseFromString(xml, 'text/xml');
        const rels = Array.from(doc.getElementsByTagName('Relationship'));
        for (const rel of rels) {
          if ((rel as Element).getAttribute('Id') === rId) {
            const t = (rel as Element).getAttribute('Target') || undefined;
            if (t) return t;
          }
        }
      } catch (e) {
        if (e instanceof Error && (/XXE|DTD|E_VALIDATION/.test(e.message))) throw e;
        /* keep scanning */
      }
    }
  }
  return undefined;
}

function findChartElement(nodeXml?: Element): Element | undefined {
  if (!nodeXml) return undefined;
  const el = nodeXml.getElementsByTagName('c:chart')?.[0] as Element | undefined;
  if (el) return el;
  const elChart = nodeXml.getElementsByTagName('chart')?.[0] as Element | undefined;
  if (elChart) return elChart;
  if (nodeXml.getElementsByTagNameNS) {
    const elNs = nodeXml.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'chart')?.[0] as Element | undefined;
    if (elNs) return elNs;
  }
  return undefined;
}

function getRId(el: Element): string | undefined {
  return (
    el.getAttribute('r:id') ||
    el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
    el.getAttribute('id') ||
    undefined
  );
}

function chartNodeIdForPart(ast: DocumentAST, parser: DocxParser, chartNum: number): Node | undefined {
  const chartPath = 'word/charts/chart' + chartNum + '.xml';
  void chartPath;
  const nodes: Node[] = [];
  for (const section of ast.sections) {
    for (const node of section.nodes) {
      if (node.type === 'chart' || node.type === 'chart3d') nodes.push(node);
    }
  }
  for (const node of nodes) {
    try {
      const chartEl = findChartElement(node.xml);
      const rId = chartEl ? getRId(chartEl) : undefined;
      if (!rId) continue;
      const target = resolveChartRIdTarget(parser, rId, ast);
      const num = target ? chartNumFromTarget(target) : undefined;
      if (num === chartNum) return node;
    } catch { /* keep scanning */ }
  }
  if (nodes.length === 1) {
    return nodes[0];
  }
  return undefined;
}

function resolveExplicitChartNode(
  op: ChartOperation,
  ast: DocumentAST,
  parser: DocxParser,
  selector: string,
  kind: string
): { chartNum: number; chartPath: string; node: Node } {
  let targetNode: Node | undefined;
  for (const section of ast.sections) {
    const found = section.nodes.find((n) => n.id === selector);
    if (found) {
      targetNode = found;
      break;
    }
  }
  if (!targetNode) {
    const valid = listChartNodeIds(ast);
    throw chartUpdateError(
      'E_VALIDATION: chart update target ' + kind + " '" + selector + "' not found.",
      valid.length > 0
        ? 'Remedy: re-read node ids via inspect_document and retry with a valid chart node id (available: ' + valid.slice(0, 8).join(', ') + (valid.length > 8 ? ' +' + (valid.length - 8) + ' more' : '') + ').'
        : "Remedy: no charts exist in this document. Insert a chart first with { action:'chart', operation:'insert', ... } and retry the update."
    );
  }
  if (targetNode.type !== 'chart' && targetNode.type !== 'chart3d') {
    throw chartUpdateError(
      'E_VALIDATION: chart update target ' + kind + " '" + selector + "' is a '" + targetNode.type + "' node, not a chart.",
      'Remedy: re-read node ids via inspect_document and retry with a chart node id (available: ' + (listChartNodeIds(ast).join(', ') || 'none — insert a chart first') + ').'
    );
  }
  const chartEl = findChartElement(targetNode.xml);
  const rId = chartEl ? getRId(chartEl) : undefined;
  if (!rId) {
    throw chartUpdateError(
      "E_VALIDATION: chart node '" + selector + "' has no resolvable chart relationship (missing c:chart r:id).",
      'Remedy: re-read the document via inspect_document and retry; if the problem persists, delete and re-insert the chart.'
    );
  }
  const target = resolveChartRIdTarget(parser, rId, ast);
  const num = target ? chartNumFromTarget(target) : undefined;
  if (target === undefined || num === undefined) {
    throw chartUpdateError(
      "E_VALIDATION: chart node '" + selector + "' references unknown relationship '" + rId + "'.",
      'Remedy: re-read the document via inspect_document and retry with a current chart node id.'
    );
  }
  const chartPath = 'word/charts/chart' + num + '.xml';
  if (!parser.getZipAdapter().hasFile(chartPath)) {
    throw chartUpdateError(
      "E_VALIDATION: chart part '" + chartPath + "' for node '" + selector + "' is missing from the package.",
      'Remedy: re-read the document via inspect_document and retry; if the problem persists, delete and re-insert the chart.'
    );
  }
  return { chartNum: num, chartPath, node: targetNode };
}

function resolveUpdateTarget(
  op: ChartOperation,
  ast: DocumentAST,
  parser: DocxParser
): { chartNum: number; chartPath: string; node?: Node } {
  const rawTarget = (op as { target?: unknown }).target;
  if (rawTarget !== undefined && rawTarget !== null) {
    if (typeof rawTarget !== 'string' || rawTarget.trim() === '') {
      throw chartUpdateError(
        'E_VALIDATION: chart update target must be a non-empty chart node id string.',
        'Remedy: retry with { action:\'chart\', operation:\'update\', target:\'<chartNodeId>\', series:[...] } using an id from inspect_document.'
      );
    }
    return resolveExplicitChartNode(op, ast, parser, rawTarget, 'target');
  }
  if (op.anchor !== undefined && op.anchor !== null) {
    if (typeof op.anchor !== 'string' || op.anchor.trim() === '') {
      throw chartUpdateError(
        'E_VALIDATION: chart update anchor must be a non-empty chart node id string.',
        'Remedy: retry with { action:\'chart\', operation:\'update\', target:\'<chartNodeId>\', series:[...] } using an id from inspect_document.'
      );
    }
    return resolveExplicitChartNode(op, ast, parser, op.anchor, 'anchor');
  }
  if (op.id !== undefined && op.id !== null) {
    if (typeof op.id !== 'string' || op.id.trim() === '') {
      throw chartUpdateError(
        'E_VALIDATION: chart update id must be a non-empty chart node id string.',
        'Remedy: retry with { action:\'chart\', operation:\'update\', target:\'<chartNodeId>\', series:[...] } using an id from inspect_document.'
      );
    }
    return resolveExplicitChartNode(op, ast, parser, op.id, 'id');
  }
  const zipAdapter = parser.getZipAdapter();
  const parts = listChartPartNumbers(zipAdapter);
  if (parts.length === 0) {
    throw chartUpdateError(
      'E_VALIDATION: chart update failed: no charts exist in this document.',
      "Remedy: insert a chart first with { action:'chart', operation:'insert', chartType:'column', title:'...', categories:[...], series:[{name:'...', values:[...]}] } and retry the update."
    );
  }
  if (parts.length > 1) {
    const valid = listChartNodeIds(ast);
    throw chartUpdateError(
      'E_VALIDATION: chart update is ambiguous: ' + parts.length + ' charts exist and no target was given.',
      valid.length > 0
        ? "Remedy: retry with an explicit chart node id, e.g. { action:'chart', operation:'update', id:'" + valid[0] + "', series:[...] } (available: " + valid.slice(0, 8).join(', ') + ').'
        : 'Remedy: re-read node ids via inspect_document and retry with an explicit chart node id.'
    );
  }
  const chartNum = parts[0];
  const chartPath = 'word/charts/chart' + chartNum + '.xml';
  const node = chartNodeIdForPart(ast, parser, chartNum);
  if (!node) {
    throw chartUpdateError(
      'E_VALIDATION: chart part ' + chartPath + ' exists but no chart AST node was found (AST/package skew).',
      'Remedy: re-read the document via inspect_document to resync the AST; if the chart was deleted from the AST, delete and re-insert the chart, then retry the update with an explicit chart node id.'
    );
  }
  return { chartNum, chartPath, node };
}

function validateUpdatePayload(op: ChartOperation): void {
  const hasMutableField =
    op.title !== undefined ||
    op.chartType !== undefined ||
    op.holeSize !== undefined ||
    op.categories !== undefined ||
    op.series !== undefined ||
    op.style !== undefined ||
    (op as { view3D?: unknown }).view3D !== undefined ||
    op.shape !== undefined;
  if (!hasMutableField) {
    throw chartUpdateError(
      'E_VALIDATION: chart update has no mutable fields to apply (need at least one of title | chartType | holeSize | categories | series | style | view3D | shape).',
      "Remedy: retry with at least one mutable field, e.g. { action:'chart', operation:'update', target:'<chartNodeId>', title:'Revenue' } or series/categories; re-read ids via inspect_document."
    );
  }
  if (op.chartType !== undefined) {
    const supported2DTypes = SUPPORTED_CHART_2D_TYPES as readonly string[];
    const supported3DTypes = SUPPORTED_CHART_TYPES as readonly string[];
    if (typeof op.chartType !== 'string' || (!supported2DTypes.includes(op.chartType) && !supported3DTypes.includes(op.chartType))) {
      throw chartUpdateError(
        'E_VALIDATION: unsupported chart type "' + op.chartType + '". Supported types: ' + SUPPORTED_CHART_2D_AND_3D_TYPES.join(', '),
        'Remedy: Use one of the supported chart types: ' + SUPPORTED_CHART_2D_AND_3D_TYPES.join(', ')
      );
    }
  }
  if (op.title !== undefined && typeof op.title !== 'string') {
    throw chartUpdateError(
      'E_VALIDATION: chart update title must be a string.',
      "Remedy: pass title as a string, e.g. { action:'chart', operation:'update', title:'Revenue' }."
    );
  }
  if (op.categories !== undefined) {
    if (!Array.isArray(op.categories) || op.categories.length === 0) {
      throw chartUpdateError(
        'E_VALIDATION: chart update categories must be a non-empty array of strings.' +
          (op.chartType === 'scatter' ? ' Scatter charts require one numeric x value per point.' : ''),
        "Remedy: pass categories like ['Q1','Q2','Q3','Q4'] or omit the field to keep existing categories; for scatter charts pass numeric x values matching every series length, e.g. ['0','5','10']."
      );
    }
    for (const c of op.categories) {
      if (typeof c !== 'string') {
        throw chartUpdateError(
          'E_VALIDATION: chart update categories must be strings.',
          "Remedy: pass categories like ['Q1','Q2'] with string entries only."
        );
      }
    }
  }
  if (op.series !== undefined) {
    if (!Array.isArray(op.series) || op.series.length === 0) {
      throw chartUpdateError(
        'E_VALIDATION: chart update series must be a non-empty array.',
        "Remedy: pass series like [{name:'ARR', values:[10,25]}] or omit the field to keep existing series."
      );
    }
    for (const s of op.series as Array<{ name?: unknown; values?: unknown }>) {
      if (!s || typeof (s as { name?: unknown }).name !== 'string') {
        throw chartUpdateError(
          'E_VALIDATION: chart update series entries need a string name.',
          "Remedy: pass series like [{name:'ARR', values:[10,25]}]."
        );
      }
      const vals = (s as { values?: unknown }).values;
      if (!Array.isArray(vals) || vals.length === 0) {
        throw chartUpdateError(
          "E_VALIDATION: chart update series '" + (s as { name?: string }).name + "' needs non-empty numeric values.",
          "Remedy: pass values like [10,25,50] with finite numbers only."
        );
      }
      for (const v of vals) {
        if (typeof v !== 'number' || !Number.isFinite(v)) {
          throw chartUpdateError(
            "E_VALIDATION: chart update series '" + (s as { name?: string }).name + "' has a non-numeric value.",
            'Remedy: pass values as finite numbers only, e.g. [10,25,50].'
          );
        }
      }
    }
  }
  if (op.shape !== undefined) {
    const allowed = ['box', 'cone', 'cylinder', 'pyramid'];
    if (!allowed.includes(op.shape)) {
      throw chartUpdateError(
        'E_VALIDATION: unsupported 3D bar shape "' + op.shape + '".',
        'Remedy: use one of box | cone | cylinder | pyramid, or omit shape to keep the existing shape.'
      );
    }
  }
}

export interface ChartStaleBodyRef {
  nodeId: string;
  preview: string;
}

export interface ChartUpdateResult {
  possiblyStaleBodyRefs: ChartStaleBodyRef[];
}

function escapeChartRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function chartParagraphTextOf(node: Node): string {
  const metaText = (node.metadata as { text?: unknown }).text;
  if (typeof metaText === 'string' && metaText.trim().length > 0) return metaText;
  try {
    const xml = node.xml as Element;
    if (!xml || typeof (xml as unknown as { getElementsByTagName?: unknown }).getElementsByTagName !== 'function') {
      return typeof metaText === 'string' ? metaText : '';
    }
    const parts: string[] = [];
    const ts = (xml as Element).getElementsByTagName('w:t');
    for (let i = 0; i < ts.length; i++) parts.push((ts[i] as Element).textContent || '');
    const joined = parts.join('');
    return joined || (typeof metaText === 'string' ? metaText : '');
  } catch {
    return typeof metaText === 'string' ? metaText : '';
  }
}

function collectPossiblyStaleBodyRefs(
  ast: DocumentAST,
  opts: { oldTitle?: string; removedCategories: string[]; removedNumbers: number[] }
): ChartStaleBodyRef[] {
  const cats = (opts.removedCategories || [])
    .map((c) => (typeof c === 'string' ? c.trim() : ''))
    .filter((c) => c.length >= 2);
  const uniqCats: string[] = [];
  const seenCats = new Set<string>();
  for (const c of cats) {
    const k = c.toLowerCase();
    if (!seenCats.has(k)) {
      seenCats.add(k);
      uniqCats.push(c);
    }
  }
  const nums = (opts.removedNumbers || []).filter((v) => typeof v === 'number' && Number.isFinite(v) && Math.abs(v) >= 100);
  const uniqNums: number[] = [...new Set(nums)];

  const rawTitle = typeof opts.oldTitle === 'string' ? opts.oldTitle.trim().replace(/\s+/g, ' ') : '';
  const titleActive = rawTitle.length >= 4;
  const titleLower = rawTitle.toLowerCase();
  const stopWords = new Set([
    'the', 'and', 'for', 'with', 'from', 'that', 'this', 'these', 'those',
    'have', 'has', 'are', 'was', 'were', 'will', 'would', 'there', 'their',
    'about', 'into', 'over', 'after', 'such',
  ]);
  const titleTokens: string[] = [];
  if (titleActive) {
    const words = rawTitle.toLowerCase().match(/[a-z0-9]+/g) || [];
    const seen = new Set<string>();
    for (const w of words) {
      if (w.length < 4 || stopWords.has(w) || seen.has(w)) continue;
      seen.add(w);
      titleTokens.push(w);
    }
  }
  if (!titleActive && uniqCats.length === 0 && uniqNums.length === 0) return [];

  const catRes = uniqCats.map((c) => new RegExp('\\b' + escapeChartRegExp(c) + '\\b', 'i'));
  const tokenRes = titleTokens.map((t) => new RegExp('\\b' + escapeChartRegExp(t) + '\\b', 'i'));

  const out: ChartStaleBodyRef[] = [];
  for (const section of ast.sections) {
    for (const node of section.nodes) {
      if (node.type !== 'paragraph') continue;
      const text = chartParagraphTextOf(node);
      if (!text || text.trim().length === 0) continue;
      const lower = text.toLowerCase();
      let hit = false;
      if (!hit && titleActive) {
        if (lower.includes(titleLower)) {
          hit = true;
        } else if (titleTokens.length >= 2) {
          let count = 0;
          for (const re of tokenRes) {
            if (re.test(text)) {
              count++;
              if (count >= 2) break;
            }
          }
          if (count >= 2) hit = true;
        }
      }
      if (!hit && catRes.length > 0) {
        for (const re of catRes) {
          if (re.test(text)) {
            hit = true;
            break;
          }
        }
      }
      if (!hit && uniqNums.length > 0) {
        const matches = text.match(/-?\$?\d[\d,]*(?:\.\d+)?/g) || [];
        for (const m of matches) {
          const cleaned = m.replace(/[$,]/g, '');
          if (cleaned === '' || cleaned === '-' || cleaned === '.') continue;
          const v = parseFloat(cleaned);
          if (!Number.isFinite(v) || Math.abs(v) < 100) continue;
          if (uniqNums.some((n) => n === v)) {
            hit = true;
            break;
          }
        }
      }
      if (hit) {
        const preview = text.length > 120 ? text.slice(0, 120) + '...' : text;
        out.push({ nodeId: node.id, preview });
      }
    }
  }
  return out;
}

/**
 * In-place update of an existing chart part: overwrites the targeted
 * word/charts/chartN.xml (series/categories/title) and refreshes the linked
 * embedded workbook, without allocating a new chart part. Fails closed with
 * E_VALIDATION when there is nothing valid to update.
 */
export function executeChartUpdate(op: ChartOperation, ast: DocumentAST, parser?: DocxParser, guide?: StyleGuideConfig): ChartUpdateResult {
  if (!parser) {
    throw chartUpdateError(
      'E_VALIDATION: chart update requires a loaded document package.',
      'Remedy: retry the update in a session backed by an open document, or re-insert the chart.'
    );
  }
  const resolved = resolveChartOp(op);
  validateUpdatePayload(resolved);
  const zipAdapter = parser.getZipAdapter();
  const targetResult = resolveUpdateTarget(resolved, ast, parser);
  const chartNum = targetResult.chartNum;
  const chartPath = targetResult.chartPath;
  const node = targetResult.node;
  if (!node) {
    throw chartUpdateError(
      'E_VALIDATION: chart part ' + chartPath + ' exists but no chart AST node was found (AST/package skew).',
      'Remedy: re-read the document via inspect_document to resync the AST; if the chart was deleted from the AST, delete and re-insert the chart, then retry the update with an explicit chart node id.'
    );
  }

  if (resolved.expectedVersion !== undefined) {
    validateVersion(node, resolved.expectedVersion);
  }

  const existingXml = zipAdapter.readText(chartPath);
  assertNoXxe(existingXml, chartPath);
  assertSupportedChartXml(existingXml, chartPath);
  const existing = parseExistingChartData(existingXml);

  const baseFamilyRequest = resolved.chartType !== undefined &&
    isBaseFamilyOf(resolved.chartType, existing.detected.chartType);
  const rawChartType = baseFamilyRequest
    ? existing.detected.chartType
    : (resolved.chartType || existing.detected.chartType);
  const supported2DTypes = SUPPORTED_CHART_2D_TYPES as readonly string[];
  const supported3DTypes = SUPPORTED_CHART_TYPES as readonly string[];
  if (!supported2DTypes.includes(rawChartType) && !supported3DTypes.includes(rawChartType)) {
    throw chartUpdateError(
      'E_VALIDATION: unsupported chart type "' + rawChartType + '". Supported types: ' + SUPPORTED_CHART_2D_AND_3D_TYPES.join(', '),
      'Remedy: Use one of the supported chart types: ' + SUPPORTED_CHART_2D_AND_3D_TYPES.join(', ')
    );
  }
  assertComboSeriesFields(resolved.series as Array<{ chartType?: unknown; axis?: unknown }> | undefined, rawChartType, chartPath);
  const scatterTypeSwitch = resolved.chartType !== undefined
    && resolved.chartType !== existing.detected.chartType && !baseFamilyRequest;
  const scatterDataSupplied = resolved.series !== undefined || resolved.categories !== undefined || scatterTypeSwitch;
  // Positional point merges only happen on data-bearing updates; a title or
  // hole-size edit on a sparse part must not refuse before any merge.
  if (rawChartType === 'scatter' && scatterDataSupplied) assertDenseScatterPoints(existingXml, chartPath);
  const use3D = is3DChartType(rawChartType);
  const holeErr = holeSizeValidationError(resolved.holeSize, rawChartType);
  if (holeErr) {
    throw chartUpdateError(
      'E_VALIDATION: ' + holeErr,
      'Remedy: use an integer holeSize from 1 to 90 for doughnut charts; omit holeSize to keep the authored doughnut hole.'
    );
  }
  const effectiveHoleSize = rawChartType === 'doughnut' ? (resolved.holeSize ?? existing.holeSize ?? 50) : undefined;

  // Duplicate requested names cannot be mapped to one stored x list per entry:
  // a repeated entry without its own xValues would inherit another entry's x
  // silently, so refuse before the mapping and before any write.
  const requestedScatterSeries = rawChartType === 'scatter'
    ? (resolved.series as Chart2DSeries[] | undefined)
    : undefined;
  if (requestedScatterSeries !== undefined && resolved.categories === undefined) {
    const seen = new Set<string>();
    const repeated = new Set<string>();
    for (const entry of requestedScatterSeries) {
      if (seen.has(entry.name)) repeated.add(entry.name);
      seen.add(entry.name);
    }
    const ambiguous = [...repeated].filter((name) =>
      requestedScatterSeries.some((entry) => entry.name === name && entry.xValues === undefined));
    if (ambiguous.length > 0) {
      const names = ambiguous.map((name) => `"${name}"`).join(', ');
      throw chartUpdateError(
        `E_VALIDATION: scatter update for chart part ${chartPath} repeats series ${ambiguous.length === 1 ? 'name' : 'names'} ${names} without categories, so a repeated entry cannot be matched to one stored x list.`,
        'Remedy: pass xValues explicitly for every repeated series entry (one finite numeric x per y value) instead of relying on the stored list.'
      );
    }
  }

  const title = resolved.title ?? existing.title;
  const categories = resolved.categories ?? existing.categories;
  const comboStoredByIndex = rawChartType === 'combo'
    ? new Map(existing.series.map((stored, index) => [stored.idx ?? index, stored]))
    : undefined;
  const requestedSeries: Chart2DSeries[] = resolved.series
    ? (resolved.series as Chart2DSeries[]).map((s, index) => {
        // A series with explicit xValues keeps them; otherwise a categories
        // argument is the new shared x list, and without one the stored
        // per-series x carries through by series name.
        const carriedX = s.xValues ?? (rawChartType === 'scatter' && resolved.categories === undefined
          ? resolveStoredScatterX(existing.series, s.name, chartPath)
          : undefined);
        const stored = existing.series[index] ?? comboStoredByIndex?.get(index);
        return {
          ...s,
          ...(carriedX !== undefined ? { xValues: carriedX } : {}),
          ...(comboStoredByIndex
            ? {
                chartType: s.chartType ?? stored?.chartType ?? 'column',
                axis: s.axis ?? stored?.axis ?? 'primary',
              }
            : {}),
        };
      })
    : (existing.series as Chart2DSeries[]);
  // Scatter resolves every series' x before merging. A categories-only update
  // moves series that still track the shared list and leaves divergent series
  // alone; a series update uses its own xValues, falling back to the stored x.
  let scatterXValues: number[][] | undefined;
  if (rawChartType === 'scatter' && scatterDataSupplied) {
    let scatterInput = requestedSeries;
    if (resolved.series === undefined && resolved.categories !== undefined) {
      const previousCategories = existing.categories.map((value) => Number(value));
      scatterInput = requestedSeries.map((s) => {
        const own = s.xValues;
        const shared = own === undefined || (own.length === previousCategories.length
          && own.every((value, i) => value === previousCategories[i]));
        return shared ? { ...s, xValues: undefined } : s;
      });
    }
    scatterXValues = assertScatterData(resolved.categories ?? existing.categories, scatterInput);
  }
  // Explicit (or already-stored) colors always win and never shift the
  // indices of the rest; only truly uncolored series pull from the guide
  // palette in order. Without a guide this reduces exactly to the legacy
  // defaultSeriesColor fallback.
  const guidePalette = cleanGuidePalette(guide);
  const overflowTaken = new Set<string>();
  const series = requestedSeries.map((s, i) => {
    const carried = (s as { color?: unknown }).color ?? existing.series[i]?.color;
    const hasExplicit = typeof carried === 'string' && carried.trim() !== '';
    const unpaired = s as Chart2DSeries & { unpairedX?: number; unpairedY?: number };
    return {
      name: s.name,
      values: s.values,
      ...(rawChartType === 'scatter'
        ? (scatterXValues?.[i] !== undefined
            ? { xValues: scatterXValues[i] }
            : s.xValues !== undefined ? { xValues: s.xValues } : {})
        : {}),
      ...(rawChartType === 'scatter' && unpaired.unpairedX !== undefined ? { unpairedX: unpaired.unpairedX } : {}),
      ...(rawChartType === 'scatter' && unpaired.unpairedY !== undefined ? { unpairedY: unpaired.unpairedY } : {}),
      ...(rawChartType === 'combo' && s.chartType !== undefined ? { chartType: s.chartType } : {}),
      ...(rawChartType === 'combo' && s.axis !== undefined ? { axis: s.axis } : {}),
      color: hasExplicit
        ? defaultSeriesColor(carried as string, i)
        : (guidePalette !== undefined
            ? guidePaletteColor(guidePalette, i, overflowTaken)
            : defaultSeriesColor(undefined, i)),
    };
  }) as Chart2DSeries[];
  // Scatter x is per series; metadata categories mirror the first series and
  // every series carries its own resolved xValues, so TOON/HTML never depend
  // on a shared fallback that could disagree with the chart part.
  const metadataCategories = rawChartType === 'scatter' && scatterXValues !== undefined && scatterXValues[0] !== undefined
    ? scatterXValues[0].map((value) => String(value))
    : categories;

  // Data-only updates patch the existing chart part in place so every node
  // the read path does not model (axes, ticks, numFmt, legend, labels,
  // trendlines, per-point formats, plot fills) survives byte-identical.
  // Changing shape (chart type) or the series count is structural and keeps
  // the historical full-regeneration path.
  const styleContext = {
    chartPath,
    chartType: rawChartType,
    seriesCount: series.length,
    is3D: use3D,
  };
  validateChartStyle(resolved.style, styleContext);

  const comboShapeChange = rawChartType === 'combo' && resolved.series !== undefined
    && resolved.series.some((rawEntry, index) => {
      const entry = rawEntry as Chart2DSeries;
      const stored = existing.series[index] ?? comboStoredByIndex?.get(index);
      if (!stored) return false;
      return (entry.chartType !== undefined && entry.chartType !== stored.chartType)
        || (entry.axis !== undefined && entry.axis !== stored.axis);
    });
  const structuralChange = scatterTypeSwitch ||
    comboShapeChange ||
    (resolved.series !== undefined && resolved.series.length !== existing.series.length);

  let updatedXml: string;
  if (!structuralChange) {
    updatedXml = mergeChartDataInPlace(existingXml, {
      chartPath,
      chartType: rawChartType,
      is3D: use3D,
      ...(resolved.title !== undefined ? { title: resolved.title } : {}),
      ...(resolved.categories !== undefined ? { categories: resolved.categories } : {}),
      ...((resolved.series !== undefined || rawChartType === 'scatter')
        && (rawChartType !== 'scatter' || scatterDataSupplied)
        ? {
            // Fill changes are opt-in: only an explicitly requested color, or a
            // color the read path actually parsed, may be patched. Otherwise the
            // stored fill (theme/gradient/pattern/no-fill) is unparsed styling and
            // must stay untouched on a data-only update.
            series: series.map((s, i) => {
              const requested = (resolved.series as Array<{ color?: unknown }> | undefined)?.[i]?.color;
              const explicit = typeof requested === 'string' && requested.trim() !== '';
              const stored = existing.series[i]?.color;
              const parsed = typeof stored === 'string' && stored.trim() !== '';
              return {
                name: s.name,
                values: s.values,
                ...(rawChartType === 'scatter' && scatterXValues !== undefined ? { xValues: scatterXValues[i] } : {}),
                ...(explicit || parsed ? { color: s.color } : {}),
              };
            }),
          }
        : {}),
      ...(use3D && resolved.view3D !== undefined ? { view3D: resolved.view3D } : {}),
      ...(use3D && resolved.shape !== undefined ? { shape: resolved.shape } : {}),
      ...(!use3D && resolved.holeSize !== undefined ? { holeSize: resolved.holeSize } : {}),
      existing: {
        title: existing.title,
        categories: existing.categories,
        series: existing.series,
      },
    });
  } else if (use3D) {
    const manipulator3D = new Chart3DManipulator(ast, parser);
    const mergedView3D: View3DProperties = { ...existing.view3D, ...(resolved.view3D || {}) };
    updatedXml = manipulator3D.generate3DChartXml(
      {
        chartType: rawChartType as Chart3DType,
        title,
        view3D: mergedView3D,
        ...(resolved.shape ? { shape: resolved.shape } : existing.shape ? { shape: existing.shape } : {}),
        categories: categories.length > 0 ? categories : existing.categories,
        series: series as unknown as Chart3DSeries[],
      },
      existing.externalDataRId
    );
  } else {
    const manipulator2D = new Chart2DManipulator(ast, parser);
    updatedXml = manipulator2D.generate2DChartXml(
      {
        chartType: rawChartType as Chart2DType,
        title,
        ...(effectiveHoleSize !== undefined ? { holeSize: effectiveHoleSize } : {}),
        categories: categories.length > 0 ? categories : existing.categories,
        series,
      },
      existing.externalDataRId
    );
  }
  if (resolved.style !== undefined) {
    updatedXml = applyChartStyle(updatedXml, resolved.style, styleContext);
    if (resolved.style.floating !== undefined) {
      const chartGraphic = findChartElement(node.xml);
      const chartRId = chartGraphic ? getRId(chartGraphic) : undefined;
      applyChartFloating(node.xml, resolved.style.floating, chartPath, chartRId);
    }
  }
  zipAdapter.writeText(chartPath, updatedXml);

  const chartRelsPath = 'word/charts/_rels/chart' + chartNum + '.xml.rels';
  const chartRelsXml = zipAdapter.readTextOrNull(chartRelsPath);
  if (chartRelsXml) assertNoXxe(chartRelsXml, chartRelsPath);
  let workbookPath: string | undefined;
  if (chartRelsXml) {
    const targetMatch = /Target="([^"]*embeddings\/[^"]*)"/.exec(chartRelsXml);
    if (targetMatch) {
      const base = targetMatch[1].split('/').pop() || '';
      if (base) workbookPath = 'word/embeddings/' + base;
    }
  }
  if (!workbookPath) {
    workbookPath = 'word/embeddings/Microsoft_Excel_Worksheet' + chartNum + '.xlsx';
  }
  // The workbook is regenerated from the parsed model, so its row count must
  // follow the data actually written into the chart cache: a series-only
  // update that carries more values than parsed categories (or a numeric /
  // unparsed category cache) must not silently drop data rows from "Edit Data".
  // A scatter edit without categories/series (title only) does not change
  // chart data; leave the workbook untouched so a malformed part cannot have
  // its x columns rewritten to point indexes.
  if (rawChartType !== 'scatter' || scatterDataSupplied) {
  const workbookCategories = categories.length > 0 ? categories : existing.categories;
  const scatterSeriesX = rawChartType === 'scatter'
    ? series.map((s) => {
        if (s.xValues !== undefined && s.xValues.length > 0) return s.xValues;
        return Array.from({ length: s.values.length }, (_, i) => {
          const numeric = Number(workbookCategories[i]);
          return Number.isFinite(numeric) ? numeric : i;
        });
      })
    : undefined;
  const workbookRowCount = Math.max(
    workbookCategories.length,
    series.reduce((max, s) => Math.max(max, s.values.length), 0),
    ...(scatterSeriesX ?? []).map((x) => x.length)
  );
  const paddedCategories = Array.from({ length: workbookRowCount }, (_, i) => workbookCategories[i] ?? '');
  const workbookBuf = generateEmbeddedWorkbook(
    series.map((s) => ({ name: s.name, values: s.values, ...(s.color ? { color: s.color } : {}) })),
    paddedCategories,
    scatterSeriesX
  );
  zipAdapter.writeBuffer(workbookPath, workbookBuf);
  }

  if (node) {
    const previousChartData = typeof node.metadata.chartData === 'object' && node.metadata.chartData !== null
      ? { ...(node.metadata.chartData as Record<string, unknown>) }
      : {};
    // Clear a prior synthetic marker only when this update supplies data; a
    // title edit alone must not turn an empty part into a hydrated-looking one.
    if (resolved.categories !== undefined || resolved.series !== undefined) {
      delete previousChartData.synthetic;
    }
    const metadataTitle = title || undefined;
    const nextMetadata: Node['metadata'] = {
      ...node.metadata,
      text: title,
      categories: metadataCategories,
      series: series as unknown as Node['metadata']['series'],
      chartType: rawChartType,
      chartData: {
        ...previousChartData,
        categories: metadataCategories,
        series,
        chartType: rawChartType,
        ...(effectiveHoleSize !== undefined ? { holeSize: effectiveHoleSize } : {}),
        ...(resolved.style !== undefined
          ? {
              style: mergeUpdatedChartStyle(previousChartData.style, resolved.style, updatedXml),
            }
          : (previousChartData.style !== undefined
              ? { style: mergeUpdatedChartStyle(previousChartData.style, {}, updatedXml) }
              : {})),
      },
    };
    // Match the reader shape exactly: a title-less chart has no title key.
    if (metadataTitle === undefined) {
      delete nextMetadata.title;
      delete (nextMetadata.chartData as Record<string, unknown>).title;
    } else {
      nextMetadata.title = metadataTitle;
      (nextMetadata.chartData as Record<string, unknown>).title = metadataTitle;
    }
    node.metadata = nextMetadata;
    node.type = use3D ? 'chart3d' : 'chart';
    node.version += 1;
    node.checksum = isChartReproducible()
      ? `chart_chk_${chartHash8(`${title}:${JSON.stringify(metadataCategories)}:${JSON.stringify(series)}:${rawChartType}`)}`
      : 'chart_chk_' + Date.now() + '_' + Math.random().toString(36).substring(2, 8);
  }

  const titleChanged = resolved.title !== undefined && resolved.title !== existing.title;
  let categoriesChanged = false;
  if (resolved.categories !== undefined) {
    const nextCats = resolved.categories;
    const prevCats = existing.categories;
    if (nextCats.length !== prevCats.length) {
      categoriesChanged = true;
    } else {
      for (let i = 0; i < nextCats.length; i++) {
        if (nextCats[i] !== prevCats[i]) {
          categoriesChanged = true;
          break;
        }
      }
    }
  }
  const seriesChanged = resolved.series !== undefined;
  const oldTitle = titleChanged ? existing.title : undefined;
  const removedCategories = categoriesChanged
    ? existing.categories.filter(
        (c) => !categories.some((n) => n.toLowerCase() === c.toLowerCase())
      )
    : [];
  const oldValues: number[] = [];
  for (const s of existing.series) {
    for (const v of s.values) oldValues.push(v);
  }
  const newValues: number[] = [];
  for (const s of series) {
    for (const v of (s.values as number[])) newValues.push(v);
  }
  const removedNumbers = seriesChanged
    ? oldValues.filter((v) => !newValues.some((n) => n === v))
    : [];
  const possiblyStaleBodyRefs = collectPossiblyStaleBodyRefs(ast, {
    ...(oldTitle !== undefined ? { oldTitle } : {}),
    removedCategories,
    removedNumbers,
  });
  return { possiblyStaleBodyRefs };
}
