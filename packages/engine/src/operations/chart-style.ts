import { chartNamespaceScopeFromXml } from '../parser/factory';
import {
  type ChartNamespaceScope,
  type Span,
  attrValue,
  chartQualifiedName,
  chartSpans,
  escapeXml,
  findChartDirectChild,
  findChartDirectChildren,
  findChartElement,
  findDirectChild,
  setAttribute,
  tagEnd,
  tagNameOf,
} from './chart-xml-spans';

/** Styling controls authored through the chart operation. */
export interface ChartLegendStyle {
  position?: 'b' | 't' | 'l' | 'r' | 'tr' | 'none';
  overlay?: boolean;
}

export interface ChartDataLabelStyle {
  show?: boolean;
  position?: 'ctr' | 'inBase' | 'inEnd' | 'outEnd' | 'bestFit' | 'b' | 't' | 'l' | 'r';
  numberFormat?: string;
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercent?: boolean;
  showLegendKey?: boolean;
}

export type ChartTickLabelPosition = 'nextTo' | 'high' | 'low' | 'none';

export interface ChartCategoryAxisStyle {
  title?: string;
  hidden?: boolean;
  tickLabelPosition?: ChartTickLabelPosition;
  majorGridlines?: boolean;
}

export interface ChartValueAxisStyle {
  title?: string;
  hidden?: boolean;
  min?: number;
  max?: number;
  numberFormat?: string;
  majorGridlines?: boolean;
  tickLabelPosition?: ChartTickLabelPosition;
}

export interface ChartAxesStyle {
  category?: ChartCategoryAxisStyle;
  value?: ChartValueAxisStyle;
}

export type ChartTrendlineType = 'linear' | 'exponential' | 'logarithmic' | 'movingAvg' | 'polynomial' | 'power' | 'none';

export interface ChartTrendlineStyle {
  type: ChartTrendlineType;
  series?: number[];
  order?: 2 | 3 | 4 | 5 | 6;
  period?: number;
  displayEquation?: boolean;
  displayRSquared?: boolean;
}

export interface ChartFloatingStyle {
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  behindText?: boolean;
  wrap?: 'none' | 'square';
}

export interface ChartStyle {
  legend?: ChartLegendStyle;
  dataLabels?: ChartDataLabelStyle;
  axes?: ChartAxesStyle;
  trendline?: ChartTrendlineStyle;
  floating?: ChartFloatingStyle;
}

export interface ChartStyleContext {
  chartPath: string;
  chartType: string;
  seriesCount: number;
  is3D: boolean;
}

const LEGEND_POSITIONS = new Set(['b', 't', 'l', 'r', 'tr', 'none']);
const LABEL_POSITIONS = new Set(['ctr', 'inBase', 'inEnd', 'outEnd', 'bestFit', 'b', 't', 'l', 'r']);
const TICK_LABEL_POSITIONS = new Set(['nextTo', 'high', 'low', 'none']);
const TRENDLINE_TYPES = new Set(['linear', 'exponential', 'logarithmic', 'movingAvg', 'polynomial', 'power', 'none']);
const PIE_TYPES = new Set(['pie', 'doughnut', 'pie3D']);
const PLOT_TAGS = [
  'barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'pieChart', 'pie3DChart',
  'doughnutChart', 'areaChart', 'scatterChart', 'bubbleChart', 'stockChart',
  'surfaceChart', 'surface3DChart', 'ofPieChart',
] as const;

function styleRefusal(context: ChartStyleContext, detail: string, remedy: string): Error {
  return Object.assign(
    new Error(`E_VALIDATION: chart style request rejected for ${context.chartPath}: ${detail}.`),
    { code: 'E_VALIDATION', retryable: false, remedy }
  );
}

function isPieFamily(chartType: string): boolean {
  return PIE_TYPES.has(chartType);
}

/** Fail-closed validation before any package mutation. */
export function validateChartStyle(style: ChartStyle | undefined, context: ChartStyleContext): void {
  if (style === undefined) return;
  if (!style || typeof style !== 'object' || Array.isArray(style)) {
    throw styleRefusal(context, 'style must be an object', 'Pass a style object such as { legend: { position: "t" } }.');
  }
  if (style.legend?.position !== undefined && !LEGEND_POSITIONS.has(style.legend.position)) {
    throw styleRefusal(context, `legend.position "${String(style.legend.position)}" is not supported`, 'Use one of b | t | l | r | tr | none.');
  }
  if (style.dataLabels?.position !== undefined && !LABEL_POSITIONS.has(style.dataLabels.position)) {
    throw styleRefusal(context, `dataLabels.position "${String(style.dataLabels.position)}" is not supported`, 'Use one of ctr | inBase | inEnd | outEnd | bestFit | b | t | l | r.');
  }
  for (const [field, value] of [
    ['dataLabels.numberFormat', style.dataLabels?.numberFormat],
    ['axes.value.numberFormat', style.axes?.value?.numberFormat],
  ] as const) {
    if (value !== undefined && (typeof value !== 'string' || value.trim() === '' || value.length > 255)) {
      throw styleRefusal(context, `${field} must be a non-empty string of at most 255 characters`, 'Use an Excel number format such as "$#,##0.00" or "0%".');
    }
  }
  if (isPieFamily(context.chartType) && style.axes !== undefined) {
    throw styleRefusal(context, `${context.chartType} charts have no value axis`, 'Omit style.axes for pie/doughnut charts; use legend and dataLabels instead.');
  }
  for (const [axisName, axis] of [
    ['category', style.axes?.category],
    ['value', style.axes?.value],
  ] as const) {
    if (axis?.title !== undefined && (typeof axis.title !== 'string' || axis.title.trim() === '' || axis.title.length > 1_000)) {
      throw styleRefusal(context, `axes.${axisName}.title must be a non-empty string of at most 1000 characters`, 'Pass the visible axis title, or omit title to keep the authored title.');
    }
    if (axis?.tickLabelPosition !== undefined && !TICK_LABEL_POSITIONS.has(axis.tickLabelPosition)) {
      throw styleRefusal(context, `axes.${axisName}.tickLabelPosition "${String(axis.tickLabelPosition)}" is not supported`, 'Use one of nextTo | high | low | none.');
    }
  }
  const valueAxis = style.axes?.value;
  for (const [field, value] of [['axes.value.min', valueAxis?.min], ['axes.value.max', valueAxis?.max]] as const) {
    if (value !== undefined && (typeof value !== 'number' || !Number.isFinite(value))) {
      throw styleRefusal(context, `${field} must be a finite number`, 'Pass a finite axis bound, or omit it to keep the authored scale.');
    }
  }
  if (valueAxis?.min !== undefined && valueAxis?.max !== undefined && valueAxis.min >= valueAxis.max) {
    throw styleRefusal(context, `axes.value.min (${valueAxis.min}) must be less than axes.value.max (${valueAxis.max})`, 'Use a min strictly below max, or omit one bound to keep the authored scale.');
  }
  if (style.floating !== undefined) {
    const floating = style.floating;
    if (floating.horizontalAlign !== undefined && !['left', 'center', 'right'].includes(floating.horizontalAlign)) {
      throw styleRefusal(context, `floating.horizontalAlign "${String(floating.horizontalAlign)}" is not supported`, 'Use left | center | right.');
    }
    if (floating.verticalAlign !== undefined && !['top', 'middle', 'bottom'].includes(floating.verticalAlign)) {
      throw styleRefusal(context, `floating.verticalAlign "${String(floating.verticalAlign)}" is not supported`, 'Use top | middle | bottom.');
    }
    if (floating.wrap !== undefined && !['none', 'square'].includes(floating.wrap)) {
      throw styleRefusal(context, `floating.wrap "${String(floating.wrap)}" is not supported`, 'Use none | square.');
    }
  }
  if (style.trendline !== undefined) {
    const trend = style.trendline;
    if (!TRENDLINE_TYPES.has(trend.type)) {
      throw styleRefusal(context, `trendline.type "${String(trend.type)}" is not supported`, 'Use one of linear | exponential | logarithmic | movingAvg | polynomial | power | none.');
    }
    if (context.is3D) {
      throw styleRefusal(context, 'trendlines are not supported on 3D chart parts', 'Use a 2D chart for trendlines, or omit style.trendline.');
    }
    if (isPieFamily(context.chartType)) {
      throw styleRefusal(context, `${context.chartType} charts cannot carry trendlines`, 'Omit style.trendline for pie/doughnut charts.');
    }
    if (trend.series !== undefined) {
      if (!Array.isArray(trend.series) || trend.series.length === 0) {
        throw styleRefusal(context, 'trendline.series must be a non-empty array of series numbers', 'Series are numbered from 1 in the order they appear: pass series like [1] or [2], or omit series to apply to every series.');
      }
      const seen = new Set<number>();
      for (const index of trend.series) {
        if (!Number.isInteger(index) || index < 1 || index > context.seriesCount) {
          throw styleRefusal(context, `trendline.series index ${String(index)} is outside 1..${context.seriesCount}`, 'Series are numbered from 1; use the number of a series that exists in the chart.');
        }
        if (seen.has(index)) {
          throw styleRefusal(context, `trendline.series index ${index} is repeated`, 'List each series number once.');
        }
        seen.add(index);
      }
    }
    if (trend.type === 'polynomial') {
      if (trend.order === undefined || !Number.isInteger(trend.order) || trend.order < 2 || trend.order > 6) {
        throw styleRefusal(context, 'polynomial trendlines require an integer order from 2 to 6', 'Pass order: 2..6.');
      }
      if (trend.period !== undefined) {
        throw styleRefusal(context, 'trendline.period is only valid for movingAvg', 'Omit period for a polynomial trendline.');
      }
    } else if (trend.type === 'movingAvg') {
      if (trend.period === undefined || !Number.isInteger(trend.period) || trend.period < 2 || trend.period > 255) {
        throw styleRefusal(context, 'movingAvg trendlines require an integer period from 2 to 255', 'Pass period: 2..255.');
      }
      if (trend.order !== undefined) {
        throw styleRefusal(context, 'trendline.order is only valid for polynomial', 'Omit order for a moving-average trendline.');
      }
    } else if (trend.order !== undefined || trend.period !== undefined) {
      throw styleRefusal(context, `order/period are not valid for ${trend.type} trendlines`, 'Use order only with polynomial and period only with movingAvg.');
    }
  }
}

function replaceSpan(source: string, span: Span, replacement: string): string {
  return source.slice(0, span.start) + replacement + source.slice(span.end);
}

function insertBeforeClose(source: string, block: string): string {
  const close = source.lastIndexOf('</');
  if (close < 0) return source + block;
  return source.slice(0, close) + block + source.slice(close);
}

function childText(elementXml: string): string | undefined {
  const match = /<(?:[A-Za-z0-9_.-]+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:[A-Za-z0-9_.-]+:)?t>/.exec(elementXml);
  if (!match) return undefined;
  return match[1]
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&');
}

const DLABELS_ORDER = [
  'dLbl', 'delete', 'numFmt', 'spPr', 'txPr', 'dLblPos', 'showLegendKey',
  'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize',
  'separator', 'showLeaderLines', 'leaderLines', 'extLst',
] as const;

/** Insert a direct child at its ECMA CT_DLbls sequence position. */
function insertDLabelChild(blockXml: string, localName: string, childXml: string): string {
  const at = DLABELS_ORDER.indexOf(localName as (typeof DLABELS_ORDER)[number]);
  if (at >= 0) {
    for (let i = at + 1; i < DLABELS_ORDER.length; i++) {
      const span = findDirectChild(blockXml, chartQualifiedName(blockXml, DLABELS_ORDER[i]));
      if (span) return blockXml.slice(0, span.start) + childXml + blockXml.slice(span.start);
    }
  }
  const gt = tagEnd(blockXml, 0);
  if (gt > 0) {
    const open = blockXml.slice(0, gt);
    const slash = open.search(/\/\s*$/);
    if (slash >= 0) return `${open.slice(0, slash)}>${childXml}</${tagNameOf(blockXml)}>`;
  }
  return insertBeforeClose(blockXml, childXml);
}

function setBoolFlag(blockXml: string, localName: string, value: boolean): string {
  const child = findDirectChild(blockXml, chartQualifiedName(blockXml, localName));
  if (child) return replaceSpan(blockXml, child, setAttribute(blockXml.slice(child.start, child.end), 'val', value ? '1' : '0'));
  const tag = chartQualifiedName(blockXml, localName);
  return insertDLabelChild(blockXml, localName, `<${tag} val="${value ? '1' : '0'}"/>`);
}

const LEGEND_ORDER = ['legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr'] as const;
const LEGEND_CHART_FOLLOWERS = ['plotVisOnly', 'dispBlanksAs', 'showDLblsOverMax', 'extLst'] as const;

function patchLegend(xml: string, chart: Span, block: string, style: ChartLegendStyle, scope: ChartNamespaceScope): string {
  if (style.position === undefined && style.overlay === undefined) return xml;
  const legend = findChartDirectChild(block, 'legend', scope);
  if (style.position === 'none') {
    if (!legend) return xml;
    const absolute = { start: chart.start + legend.start, end: chart.start + legend.end };
    return replaceSpan(xml, absolute, '');
  }
  if (legend) {
    let legendXml = block.slice(legend.start, legend.end);
    if (style.position !== undefined) {
      const pos = findDirectChild(legendXml, chartQualifiedName(legendXml, 'legendPos'));
      if (pos) {
        legendXml = replaceSpan(legendXml, pos, setAttribute(legendXml.slice(pos.start, pos.end), 'val', style.position));
      } else {
        const tag = chartQualifiedName(legendXml, 'legendPos');
        legendXml = insertOrderedChild(legendXml, LEGEND_ORDER, 'legendPos', `<${tag} val="${style.position}"/>`);
      }
    }
    if (style.overlay !== undefined) {
      const overlay = findDirectChild(legendXml, chartQualifiedName(legendXml, 'overlay'));
      if (overlay) {
        legendXml = replaceSpan(legendXml, overlay, setAttribute(legendXml.slice(overlay.start, overlay.end), 'val', style.overlay ? '1' : '0'));
      } else {
        const tag = chartQualifiedName(legendXml, 'overlay');
        legendXml = insertOrderedChild(legendXml, LEGEND_ORDER, 'overlay', `<${tag} val="${style.overlay ? '1' : '0'}"/>`);
      }
    }
    return replaceSpan(xml, { start: chart.start + legend.start, end: chart.start + legend.end }, legendXml);
  }
  const tag = chartQualifiedName(block, 'legend');
  const pos = chartQualifiedName(block, 'legendPos');
  const overlay = chartQualifiedName(block, 'overlay');
  const legendXml = `<${tag}><${pos} val="${style.position ?? 'r'}"/><${overlay} val="${style.overlay ? '1' : '0'}"/></${tag}>`;
  const plotArea = findChartDirectChild(block, 'plotArea', scope);
  let at = chart.start + block.lastIndexOf('</');
  if (plotArea) {
    at = chart.start + plotArea.end;
  } else {
    for (const follower of LEGEND_CHART_FOLLOWERS) {
      const span = findChartDirectChild(block, follower, scope);
      if (span) {
        at = chart.start + span.start;
        break;
      }
    }
  }
  return xml.slice(0, at) + legendXml + xml.slice(at);
}

function effectiveValueAxis(block: string, scope: ChartNamespaceScope): Span | null {
  const axes = chartSpans(block, 'valAx', scope);
  if (axes.length === 0) return null;
  if (axes.length === 1) return axes[0];
  for (const axis of axes) {
    const axisXml = block.slice(axis.start, axis.end);
    const pos = findDirectChild(axisXml, chartQualifiedName(axisXml, 'axPos'));
    if (pos) {
      const val = attrValue(axisXml.slice(pos.start, pos.end), 'val');
      if (val === 'l' || val === 'r') return axis;
    }
  }
  return axes[axes.length - 1];
}

function allPlots(block: string, scope: ChartNamespaceScope): Array<{ span: Span; tag: string }> {
  const plotArea = findChartDirectChild(block, 'plotArea', scope);
  if (!plotArea) return [];
  const areaXml = block.slice(plotArea.start, plotArea.end);
  const plots: Array<{ span: Span; tag: string }> = [];
  for (const tag of PLOT_TAGS) {
    for (const span of findChartDirectChildren(areaXml, tag, scope)) {
      plots.push({
        span: { start: plotArea.start + span.start, end: plotArea.start + span.end },
        tag,
      });
    }
  }
  return plots.sort((a, b) => a.span.start - b.span.start);
}

function patchDataLabels(xml: string, chart: Span, block: string, style: ChartDataLabelStyle, scope: ChartNamespaceScope): string {
  const hasRequest = style.show !== undefined
    || style.position !== undefined
    || style.numberFormat !== undefined
    || style.showValue !== undefined
    || style.showCategoryName !== undefined
    || style.showSeriesName !== undefined
    || style.showPercent !== undefined
    || style.showLegendKey !== undefined;
  if (!hasRequest) return xml;
  const plots = allPlots(block, scope);
  if (plots.length === 0) return xml;
  let patchedBlock = block;
  for (const plot of [...plots].reverse()) {
    let plotXml = patchedBlock.slice(plot.span.start, plot.span.end);
    const dLbls = findDirectChild(plotXml, chartQualifiedName(plotXml, 'dLbls'));
    const requestedFlags: Array<[string, boolean]> = [];
    if (style.show === false) {
      for (const flag of ['showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']) {
        requestedFlags.push([flag, false]);
      }
    } else if (style.show === true) {
      requestedFlags.push(['showVal', true]);
      for (const flag of ['showLegendKey', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize']) {
        requestedFlags.push([flag, false]);
      }
    }
    if (style.showValue !== undefined) requestedFlags.push(['showVal', style.showValue]);
    if (style.showCategoryName !== undefined) requestedFlags.push(['showCatName', style.showCategoryName]);
    if (style.showSeriesName !== undefined) requestedFlags.push(['showSerName', style.showSeriesName]);
    if (style.showPercent !== undefined) requestedFlags.push(['showPercent', style.showPercent]);
    if (style.showLegendKey !== undefined) requestedFlags.push(['showLegendKey', style.showLegendKey]);

    if (dLbls) {
      let labelsXml = plotXml.slice(dLbls.start, dLbls.end);
      for (const [flag, value] of requestedFlags) labelsXml = setBoolFlag(labelsXml, flag, value);
      if (style.show === true && findDirectChild(labelsXml, chartQualifiedName(labelsXml, 'delete'))) {
        labelsXml = setBoolFlag(labelsXml, 'delete', false);
      }
      if (style.show === false) labelsXml = setBoolFlag(labelsXml, 'delete', true);
      if (style.position !== undefined) {
        const pos = findDirectChild(labelsXml, chartQualifiedName(labelsXml, 'dLblPos'));
        if (pos) {
          labelsXml = replaceSpan(labelsXml, pos, setAttribute(labelsXml.slice(pos.start, pos.end), 'val', style.position));
        } else {
          const tag = chartQualifiedName(labelsXml, 'dLblPos');
          labelsXml = insertDLabelChild(labelsXml, 'dLblPos', `<${tag} val="${style.position}"/>`);
        }
      }
      if (style.numberFormat !== undefined) {
        const numFmt = findDirectChild(labelsXml, chartQualifiedName(labelsXml, 'numFmt'));
        if (numFmt) {
          let patched = setAttribute(labelsXml.slice(numFmt.start, numFmt.end), 'formatCode', escapeXml(style.numberFormat));
          patched = setAttribute(patched, 'sourceLinked', '0');
          labelsXml = replaceSpan(labelsXml, numFmt, patched);
        } else {
          const tag = chartQualifiedName(labelsXml, 'numFmt');
          labelsXml = insertDLabelChild(labelsXml, 'numFmt', `<${tag} formatCode="${escapeXml(style.numberFormat)}" sourceLinked="0"/>`);
        }
      }
      plotXml = replaceSpan(plotXml, dLbls, labelsXml);
    } else {
      const tag = chartQualifiedName(plotXml, 'dLbls');
      const numFmtTag = chartQualifiedName(plotXml, 'numFmt');
      const byTag = new Map<string, string>();
      if (style.show === false) byTag.set('delete', `<${chartQualifiedName(plotXml, 'delete')} val="1"/>`);
      if (style.numberFormat !== undefined) byTag.set('numFmt', `<${numFmtTag} formatCode="${escapeXml(style.numberFormat)}" sourceLinked="0"/>`);
      if (style.position !== undefined) byTag.set('dLblPos', `<${chartQualifiedName(plotXml, 'dLblPos')} val="${style.position}"/>`);
      for (const [flag, value] of requestedFlags) byTag.set(flag, `<${chartQualifiedName(plotXml, flag)} val="${value ? '1' : '0'}"/>`);
      const parts = DLABELS_ORDER.filter((name) => byTag.has(name)).map((name) => byTag.get(name)!);
      const labelsXml = `<${tag}>${parts.join('')}</${tag}>`;
      const sers = chartSpans(plotXml, 'ser', scope);
      const anchor = sers.length > 0
        ? { start: sers[sers.length - 1].end, end: sers[sers.length - 1].end }
        : null;
      if (anchor) {
        plotXml = plotXml.slice(0, anchor.start) + labelsXml + plotXml.slice(anchor.start);
      } else {
        plotXml = insertBeforeClose(plotXml, labelsXml);
      }
    }
    patchedBlock = replaceSpan(patchedBlock, plot.span, plotXml);
  }
  return replaceSpan(xml, { start: chart.start, end: chart.end }, patchedBlock);
}


const AXIS_ORDER = [
  'axId', 'scaling', 'delete', 'axPos', 'majorGridlines', 'minorGridlines',
  'title', 'numFmt', 'majorTickMark', 'minorTickMark', 'tickLblPos',
  'spPr', 'txPr', 'crossAx', 'crosses', 'crossesAt', 'crossBetween',
  'auto', 'lblAlgn', 'lblOffset', 'tickLblSkip', 'tickMarkSkip',
  'noMultiLvlLbl', 'majorUnit', 'minorUnit', 'dispUnits', 'extLst',
] as const;

const SCALING_ORDER = ['logBase', 'orientation', 'max', 'min'] as const;

function insertOrderedChild(
  blockXml: string,
  order: readonly string[],
  localName: string,
  childXml: string
): string {
  const at = order.indexOf(localName);
  if (at >= 0) {
    for (let i = at + 1; i < order.length; i++) {
      const span = findDirectChild(blockXml, chartQualifiedName(blockXml, order[i]));
      if (span) return blockXml.slice(0, span.start) + childXml + blockXml.slice(span.start);
    }
  }
  const gt = tagEnd(blockXml, 0);
  if (gt > 0) {
    const open = blockXml.slice(0, gt);
    const slash = open.search(/\/\s*$/);
    if (slash >= 0) return `${open.slice(0, slash)}>${childXml}</${tagNameOf(blockXml)}>`;
  }
  return insertBeforeClose(blockXml, childXml);
}

function setAxisChild(
  axisXml: string,
  localName: string,
  value: string,
  order: readonly string[] = AXIS_ORDER
): string {
  const tag = chartQualifiedName(axisXml, localName);
  const child = findDirectChild(axisXml, tag);
  if (child) return replaceSpan(axisXml, child, setAttribute(axisXml.slice(child.start, child.end), 'val', value));
  return insertOrderedChild(axisXml, order, localName, `<${tag} val="${value}"/>`);
}

function removeDirectChildren(blockXml: string, localName: string): string {
  let out = blockXml;
  const tag = chartQualifiedName(blockXml, localName);
  for (;;) {
    const child = findDirectChild(out, tag);
    if (!child) return out;
    out = out.slice(0, child.start) + out.slice(child.end);
  }
}

function patchScalingBounds(axisXml: string, min: number | undefined, max: number | undefined): string {
  if (min === undefined && max === undefined) return axisXml;
  let out = axisXml;
  let scaling = findDirectChild(out, chartQualifiedName(out, 'scaling'));
  if (!scaling) {
    const tag = chartQualifiedName(out, 'scaling');
    const orientation = chartQualifiedName(out, 'orientation');
    const block = `<${tag}><${orientation} val="minMax"/></${tag}>`;
    out = insertOrderedChild(out, AXIS_ORDER, 'scaling', block);
    scaling = findDirectChild(out, chartQualifiedName(out, 'scaling'));
    if (!scaling) return out;
  }
  let scalingXml = out.slice(scaling.start, scaling.end);
  if (max !== undefined) {
    const tag = chartQualifiedName(scalingXml, 'max');
    const child = findDirectChild(scalingXml, tag);
    if (child) scalingXml = replaceSpan(scalingXml, child, setAttribute(scalingXml.slice(child.start, child.end), 'val', String(max)));
    else scalingXml = insertOrderedChild(scalingXml, SCALING_ORDER, 'max', `<${tag} val="${max}"/>`);
  }
  if (min !== undefined) {
    const tag = chartQualifiedName(scalingXml, 'min');
    const child = findDirectChild(scalingXml, tag);
    if (child) scalingXml = replaceSpan(scalingXml, child, setAttribute(scalingXml.slice(child.start, child.end), 'val', String(min)));
    else scalingXml = insertOrderedChild(scalingXml, SCALING_ORDER, 'min', `<${tag} val="${min}"/>`);
  }
  return replaceSpan(out, scaling, scalingXml);
}

function axisTitleBlock(axisXml: string, title: string): string {
  const c = (localName: string): string => chartQualifiedName(axisXml, localName);
  const xmlSpace = /^\s|\s$/.test(title) ? ' xml:space="preserve"' : '';
  return `<${c('title')}><${c('tx')}><${c('rich')}><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1000"/><a:t${xmlSpace}>${escapeXml(title)}</a:t></a:r></a:p></${c('rich')}></${c('tx')}><${c('overlay')} val="0"/></${c('title')}>`;
}

function patchAxis(
  axisXml: string,
  style: {
    title?: string;
    hidden?: boolean;
    tickLabelPosition?: ChartTickLabelPosition;
    majorGridlines?: boolean;
    min?: number;
    max?: number;
    numberFormat?: string;
  }
): string {
  let out = axisXml;
  if (style.hidden !== undefined) out = setAxisChild(out, 'delete', style.hidden ? '1' : '0');
  if (style.tickLabelPosition !== undefined) out = setAxisChild(out, 'tickLblPos', style.tickLabelPosition);
  if (style.majorGridlines !== undefined) {
    if (style.majorGridlines) {
      if (!findDirectChild(out, chartQualifiedName(out, 'majorGridlines'))) {
        const tag = chartQualifiedName(out, 'majorGridlines');
        out = insertOrderedChild(out, AXIS_ORDER, 'majorGridlines', `<${tag}/>`);
      }
    } else {
      out = removeDirectChildren(out, 'majorGridlines');
    }
  }
  if (style.title !== undefined) {
    const tag = chartQualifiedName(out, 'title');
    const existing = findDirectChild(out, tag);
    const block = axisTitleBlock(out, style.title);
    out = existing ? replaceSpan(out, existing, block) : insertOrderedChild(out, AXIS_ORDER, 'title', block);
  }
  if (style.numberFormat !== undefined) {
    const tag = chartQualifiedName(out, 'numFmt');
    const existing = findDirectChild(out, tag);
    if (existing) {
      let patched = setAttribute(out.slice(existing.start, existing.end), 'formatCode', escapeXml(style.numberFormat));
      patched = setAttribute(patched, 'sourceLinked', '0');
      out = replaceSpan(out, existing, patched);
    } else {
      out = insertOrderedChild(out, AXIS_ORDER, 'numFmt', `<${tag} formatCode="${escapeXml(style.numberFormat)}" sourceLinked="0"/>`);
    }
  }
  return patchScalingBounds(out, style.min, style.max);
}

function axisSpan(block: string, kind: 'category' | 'value', context: ChartStyleContext, scope: ChartNamespaceScope): Span | null {
  if (kind === 'value') return effectiveValueAxis(block, scope);
  const cat = chartSpans(block, 'catAx', scope)[0];
  if (cat) return cat;
  if (context.chartType === 'scatter') {
    for (const axis of chartSpans(block, 'valAx', scope)) {
      const axisXml = block.slice(axis.start, axis.end);
      const pos = findDirectChild(axisXml, chartQualifiedName(axisXml, 'axPos'));
      if (pos && attrValue(axisXml.slice(pos.start, pos.end), 'val') === 'b') return axis;
    }
  }
  return null;
}

function axisWithId(block: string, kind: string, id: string, scope: ChartNamespaceScope): Span | null {
  for (const axis of chartSpans(block, kind, scope)) {
    const axisXml = block.slice(axis.start, axis.end);
    const axId = findDirectChild(axisXml, chartQualifiedName(axisXml, 'axId'));
    if (axId && attrValue(axisXml.slice(axId.start, axId.end), 'val') === id) return axis;
  }
  return null;
}

function patchAxes(
  xml: string,
  style: ChartAxesStyle,
  context: ChartStyleContext,
  scope: ChartNamespaceScope
): string {
  let out = xml;
  for (const kind of ['category', 'value'] as const) {
    const axisStyle = style[kind];
    if (axisStyle === undefined) continue;
    const chart = findChartElement(out, 'chart', scope);
    if (!chart) return out;
    const block = out.slice(chart.start, chart.end);
    const axis = axisSpan(block, kind, context, scope);
    if (!axis) {
      throw styleRefusal(
        context,
        kind === 'category' ? 'chart part has no category axis to style' : 'chart part has no value axis to style',
        kind === 'category' && context.chartType === 'scatter'
          ? 'Scatter charts expose their numeric x axis as the category axis; re-read the part and retry.'
          : 'Re-read the chart via inspect_document; if malformed, delete and re-insert it.'
      );
    }
    const axisXml = patchAxis(block.slice(axis.start, axis.end), axisStyle);
    out = replaceSpan(out, { start: chart.start + axis.start, end: chart.start + axis.end }, axisXml);
    const numberFormat = kind === 'value' ? style.value?.numberFormat : undefined;
    if (kind === 'value' && context.chartType === 'combo' && numberFormat !== undefined) {
      const updatedChart = findChartElement(out, 'chart', scope);
      if (!updatedChart) return out;
      const updatedBlock = out.slice(updatedChart.start, updatedChart.end);
      const secondary = axisWithId(updatedBlock, 'valAx', '2002', scope);
      if (secondary) {
        const secondaryXml = patchAxis(updatedBlock.slice(secondary.start, secondary.end), { numberFormat });
        out = replaceSpan(out, { start: updatedChart.start + secondary.start, end: updatedChart.start + secondary.end }, secondaryXml);
      }
    }
  }
  return out;
}

const TRENDLINE_AFTER_TAGS = ['errBars', 'cat', 'val', 'xVal', 'yVal', 'shape', 'extLst'] as const;

function trendlineBlock(seriesXml: string, style: ChartTrendlineStyle): string {
  const c = (localName: string): string => chartQualifiedName(seriesXml, localName);
  const parts = [`<${c('trendlineType')} val="${style.type}"/>`];
  if (style.order !== undefined) parts.push(`<${c('order')} val="${style.order}"/>`);
  if (style.period !== undefined) parts.push(`<${c('period')} val="${style.period}"/>`);
  if (style.displayRSquared !== undefined) parts.push(`<${c('dispRSqr')} val="${style.displayRSquared ? '1' : '0'}"/>`);
  if (style.displayEquation !== undefined) parts.push(`<${c('dispEq')} val="${style.displayEquation ? '1' : '0'}"/>`);
  return `<${c('trendline')}>${parts.join('')}</${c('trendline')}>`;
}

function patchTrendlines(
  xml: string,
  style: ChartTrendlineStyle,
  context: ChartStyleContext,
  scope: ChartNamespaceScope
): string {
  const chart = findChartElement(xml, 'chart', scope);
  if (!chart) return xml;
  let block = xml.slice(chart.start, chart.end);
  const sers = chartSpans(block, 'ser', scope);
  if (sers.length === 0) {
    throw styleRefusal(context, 'chart has no series to attach a trendline to', 'Add series data first, or omit style.trendline.');
  }
  // Callers count series from 1; the chart part's c:idx values count from 0.
  const selected = (style.series ?? Array.from({ length: sers.length }, (_, i) => i + 1)).map((n) => n - 1);
  for (const index of selected) {
    if (index < 0 || index >= sers.length) {
      throw styleRefusal(context, `trendline.series index ${index + 1} is outside 1..${sers.length}`, 'Series are numbered from 1; use the number of a series that exists in the chart part.');
    }
  }
  const selectedSet = new Set(selected);
  for (let i = sers.length - 1; i >= 0; i--) {
    let serXml = block.slice(sers[i].start, sers[i].end);
    const idxSpan = findDirectChild(serXml, chartQualifiedName(serXml, 'idx'));
    const seriesIndex = idxSpan ? Number(attrValue(serXml.slice(idxSpan.start, idxSpan.end), 'val')) : i;
    if (selectedSet.has(seriesIndex)) {
      const existing = (() => {
        const spans: Span[] = [];
        let cursor = 0;
        for (;;) {
          const span = findDirectChild(serXml.slice(cursor), chartQualifiedName(serXml, 'trendline'));
          if (!span) break;
          const absolute = { start: cursor + span.start, end: cursor + span.end };
          spans.push(absolute);
          cursor = absolute.end;
        }
        return spans;
      })();
      for (const span of [...existing].reverse()) serXml = serXml.slice(0, span.start) + serXml.slice(span.end);
      if (style.type !== 'none') {
        const child = trendlineBlock(serXml, style);
        let at = serXml.lastIndexOf('</');
        for (const tag of TRENDLINE_AFTER_TAGS) {
          const span = findDirectChild(serXml, chartQualifiedName(serXml, tag));
          if (span) {
            at = span.start;
            break;
          }
        }
        serXml = serXml.slice(0, at) + child + serXml.slice(at);
      }
    }
    block = replaceSpan(block, sers[i], serXml);
  }
  return replaceSpan(xml, { start: chart.start, end: chart.end }, block);
}


function readAxisStyle(axisXml: string): Partial<ChartValueAxisStyle> {
  const out: Partial<ChartValueAxisStyle> = {};
  const c = (localName: string): string => chartQualifiedName(axisXml, localName);
  const title = findDirectChild(axisXml, c('title'));
  if (title) {
    const text = childText(axisXml.slice(title.start, title.end));
    if (text !== undefined && text.trim() !== '') out.title = text;
  }
  const del = findDirectChild(axisXml, c('delete'));
  if (del) {
    const val = attrValue(axisXml.slice(del.start, del.end), 'val');
    if (val !== undefined) out.hidden = val === '1' || val === 'true';
  }
  const tick = findDirectChild(axisXml, c('tickLblPos'));
  if (tick) {
    const val = attrValue(axisXml.slice(tick.start, tick.end), 'val');
    if (val && TICK_LABEL_POSITIONS.has(val)) out.tickLabelPosition = val as ChartTickLabelPosition;
  }
  if (findDirectChild(axisXml, c('majorGridlines'))) out.majorGridlines = true;
  else out.majorGridlines = false;
  const numFmt = findDirectChild(axisXml, c('numFmt'));
  if (numFmt) {
    const code = attrValue(axisXml.slice(numFmt.start, numFmt.end), 'formatCode');
    if (code) out.numberFormat = code;
  }
  const scaling = findDirectChild(axisXml, c('scaling'));
  if (scaling) {
    const scalingXml = axisXml.slice(scaling.start, scaling.end);
    const max = findDirectChild(scalingXml, chartQualifiedName(scalingXml, 'max'));
    const min = findDirectChild(scalingXml, chartQualifiedName(scalingXml, 'min'));
    if (max) {
      const val = attrValue(scalingXml.slice(max.start, max.end), 'val');
      if (val !== undefined && Number.isFinite(Number(val))) out.max = Number(val);
    }
    if (min) {
      const val = attrValue(scalingXml.slice(min.start, min.end), 'val');
      if (val !== undefined && Number.isFinite(Number(val))) out.min = Number(val);
    }
  }
  return out;
}

function readTrendlineStyle(block: string, scope: ChartNamespaceScope): ChartTrendlineStyle | undefined {
  let result: ChartTrendlineStyle | undefined;
  let ambiguousType = false;
  const seriesIndexes = new Set<number>();
  for (const plot of allPlots(block, scope)) {
    const plotXml = block.slice(plot.span.start, plot.span.end);
    for (const ser of chartSpans(plotXml, 'ser', scope)) {
      const serXml = plotXml.slice(ser.start, ser.end);
      const trend = findDirectChild(serXml, chartQualifiedName(serXml, 'trendline'));
      if (!trend) continue;
      const trendXml = serXml.slice(trend.start, trend.end);
      const typeEl = findDirectChild(trendXml, chartQualifiedName(trendXml, 'trendlineType'));
      if (!typeEl) continue;
      const type = attrValue(trendXml.slice(typeEl.start, typeEl.end), 'val');
      if (!type || !TRENDLINE_TYPES.has(type)) continue;
      if (result === undefined) {
        result = { type: type as ChartTrendlineType, series: [] };
        const order = findDirectChild(trendXml, chartQualifiedName(trendXml, 'order'));
        const period = findDirectChild(trendXml, chartQualifiedName(trendXml, 'period'));
        const rsq = findDirectChild(trendXml, chartQualifiedName(trendXml, 'dispRSqr'));
        const eq = findDirectChild(trendXml, chartQualifiedName(trendXml, 'dispEq'));
        if (order) {
          const val = Number(attrValue(trendXml.slice(order.start, order.end), 'val'));
          if (Number.isInteger(val) && val >= 2 && val <= 6) result.order = val as ChartTrendlineStyle['order'];
        }
        if (period) {
          const val = Number(attrValue(trendXml.slice(period.start, period.end), 'val'));
          if (Number.isInteger(val) && val >= 2 && val <= 255) result.period = val;
        }
        if (rsq) result.displayRSquared = attrValue(trendXml.slice(rsq.start, rsq.end), 'val') === '1';
        if (eq) result.displayEquation = attrValue(trendXml.slice(eq.start, eq.end), 'val') === '1';
      } else if (result.type !== type) {
        ambiguousType = true;
        continue;
      }
      const idx = findDirectChild(serXml, chartQualifiedName(serXml, 'idx'));
      const index = idx ? Number(attrValue(serXml.slice(idx.start, idx.end), 'val')) : Number.NaN;
      if (Number.isInteger(index) && index >= 0) seriesIndexes.add(index + 1);
    }
  }
  if (!result || ambiguousType) return undefined;
  result.series = [...seriesIndexes].sort((a, b) => a - b);
  return result;
}

function readDataLabels(plotXml: string): ChartDataLabelStyle | undefined {
  const labels = findDirectChild(plotXml, chartQualifiedName(plotXml, 'dLbls'));
  if (!labels) return undefined;
  const labelsXml = plotXml.slice(labels.start, labels.end);
  const readFlag = (name: string): boolean | undefined => {
    const child = findDirectChild(labelsXml, chartQualifiedName(labelsXml, name));
    if (!child) return undefined;
    const val = attrValue(labelsXml.slice(child.start, child.end), 'val');
    return val === undefined ? undefined : val === '1' || val === 'true';
  };
  const dataLabels: ChartDataLabelStyle = {};
  const showVal = readFlag('showVal');
  const showCat = readFlag('showCatName');
  const showSer = readFlag('showSerName');
  const showPct = readFlag('showPercent');
  const showKey = readFlag('showLegendKey');
  if (showVal === true || showCat === true || showSer === true || showPct === true || showKey === true) dataLabels.show = true;
  if (showVal === false && showCat === false && showSer === false && showPct === false && showKey === false) dataLabels.show = false;
  if (showVal !== undefined) dataLabels.showValue = showVal;
  if (showCat !== undefined) dataLabels.showCategoryName = showCat;
  if (showSer !== undefined) dataLabels.showSeriesName = showSer;
  if (showPct !== undefined) dataLabels.showPercent = showPct;
  if (showKey !== undefined) dataLabels.showLegendKey = showKey;
  const pos = findDirectChild(labelsXml, chartQualifiedName(labelsXml, 'dLblPos'));
  if (pos) {
    const val = attrValue(labelsXml.slice(pos.start, pos.end), 'val');
    if (val && LABEL_POSITIONS.has(val)) dataLabels.position = val as ChartDataLabelStyle['position'];
  }
  const numFmt = findDirectChild(labelsXml, chartQualifiedName(labelsXml, 'numFmt'));
  if (numFmt) {
    const code = attrValue(labelsXml.slice(numFmt.start, numFmt.end), 'formatCode');
    if (code) dataLabels.numberFormat = code;
  }
  return Object.keys(dataLabels).length > 0 ? dataLabels : undefined;
}

function commonDataLabels(styles: Array<ChartDataLabelStyle | undefined>): ChartDataLabelStyle | undefined {
  if (styles.length === 0 || styles.some((entry) => entry === undefined)) return undefined;
  const out: ChartDataLabelStyle = {};
  const keys = Object.keys(styles[0]!) as Array<keyof ChartDataLabelStyle>;
  for (const key of keys) {
    const value = styles[0]![key];
    if (styles.every((entry) => entry![key] === value)) {
      (out as Record<string, unknown>)[key] = value;
    }
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** Pure string transform: validate before calling; throws before returning on unsafe input. */
export function applyChartStyle(xml: string, style: ChartStyle | undefined, context: ChartStyleContext): string {
  if (style === undefined) return xml;
  validateChartStyle(style, context);
  const scope = chartNamespaceScopeFromXml(xml);
  const chart = findChartElement(xml, 'chart', scope);
  if (!chart) {
    throw styleRefusal(context, 'chart part has no <chart> root child', 'Re-read the chart via inspect_document; if malformed, delete and re-insert it.');
  }
  let out = xml;
  let block = out.slice(chart.start, chart.end);
  if (style.legend !== undefined) {
    out = patchLegend(out, chart, block, style.legend, scope);
    block = out.slice(chart.start, chart.end);
  }
  if (style.axes !== undefined) {
    out = patchAxes(out, style.axes, context, scope);
  }
  if (style.trendline !== undefined) {
    out = patchTrendlines(out, style.trendline, context, scope);
  }
  if (style.dataLabels !== undefined) {
    const current = findChartElement(out, 'chart', scope);
    if (current) {
      block = out.slice(current.start, current.end);
      out = patchDataLabels(out, current, block, style.dataLabels, scope);
    }
  }
  return out;
}

/** Generated legacy defaults are not authored styling; do not hydrate a style from them alone. */
function isLegacyDefaultStyle(style: ChartStyle): boolean {
  for (const key of Object.keys(style)) {
    if (key !== 'legend' && key !== 'axes') return false;
  }
  if (style.legend) {
    for (const [key, value] of Object.entries(style.legend)) {
      if (key === 'position' && value === 'r') continue;
      if (key === 'overlay' && value === false) continue;
      return false;
    }
  }
  const category = style.axes?.category as Record<string, unknown> | undefined;
  if (category) {
    for (const [key, value] of Object.entries(category)) {
      if (key === 'hidden' && value === false) continue;
      if (key === 'tickLabelPosition' && value === 'nextTo') continue;
      if (key === 'numberFormat' && value === 'General') continue;
      if (key === 'majorGridlines' && value === false) continue;
      return false;
    }
  }
  const valueAxis = style.axes?.value;
  if (valueAxis) {
    for (const [key, value] of Object.entries(valueAxis)) {
      if (key === 'hidden' && value === false) continue;
      if (key === 'tickLabelPosition' && value === 'nextTo') continue;
      if (key === 'majorGridlines' && value === true) continue;
      if (key === 'numberFormat' && value === 'General') continue;
      return false;
    }
  }
  return true;
}

/** Reads only fields that are actually present; returns undefined for an unstyled part. */
export function readChartStyle(xml: string): ChartStyle | undefined {
  if (typeof xml !== 'string' || xml.length === 0) return undefined;
  let scope: ChartNamespaceScope;
  try {
    scope = chartNamespaceScopeFromXml(xml);
  } catch {
    return undefined;
  }
  const chart = findChartElement(xml, 'chart', scope);
  if (!chart) return undefined;
  const block = xml.slice(chart.start, chart.end);
  const style: ChartStyle = {};
  const legend = findChartDirectChild(block, 'legend', scope);
  if (legend) {
    const legendXml = block.slice(legend.start, legend.end);
    const pos = findDirectChild(legendXml, chartQualifiedName(legendXml, 'legendPos'));
    if (pos) {
      const val = attrValue(legendXml.slice(pos.start, pos.end), 'val');
      if (val && LEGEND_POSITIONS.has(val) && val !== 'none') style.legend = { ...(style.legend ?? {}), position: val as ChartLegendStyle['position'] };
    }
    const overlay = findDirectChild(legendXml, chartQualifiedName(legendXml, 'overlay'));
    if (overlay) {
      const val = attrValue(legendXml.slice(overlay.start, overlay.end), 'val');
      if (val !== undefined) style.legend = { ...(style.legend ?? {}), overlay: val === '1' || val === 'true' };
    }
  }
  const axes: ChartAxesStyle = {};
  let categoryAxis: Span | undefined = chartSpans(block, 'catAx', scope)[0];
  if (!categoryAxis) {
    const valAxes = chartSpans(block, 'valAx', scope);
    if (valAxes.length > 1) {
      categoryAxis = valAxes.find((span) => {
        const axisXml = block.slice(span.start, span.end);
        const pos = findDirectChild(axisXml, chartQualifiedName(axisXml, 'axPos'));
        return pos !== null && attrValue(axisXml.slice(pos.start, pos.end), 'val') === 'b';
      });
    }
  }
  if (categoryAxis) {
    const parsed = readAxisStyle(block.slice(categoryAxis.start, categoryAxis.end));
    if (Object.keys(parsed).length > 0) axes.category = parsed as ChartCategoryAxisStyle;
  }
  const valueAxis = effectiveValueAxis(block, scope);
  if (valueAxis) {
    const parsed = readAxisStyle(block.slice(valueAxis.start, valueAxis.end));
    if (Object.keys(parsed).length > 0) axes.value = parsed;
  }
  if (axes.category !== undefined || axes.value !== undefined) style.axes = axes;
  const plots = allPlots(block, scope);
  if (plots.length > 0) {
    const labels = commonDataLabels(plots.map((plot) => readDataLabels(block.slice(plot.span.start, plot.span.end))));
    if (labels) style.dataLabels = labels;
    const trendline = readTrendlineStyle(block, scope);
    if (trendline) style.trendline = trendline;
  }
  if (Object.keys(style).length === 0) return undefined;
  return isLegacyDefaultStyle(style) ? undefined : style;
}

/** Exported for tests that need to assert a title was not left behind. */
export function chartElementText(xml: string): string | undefined {
  return childText(xml);
}

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

function directChildByLocal(parent: Element, localName: string): Element | undefined {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i] as Element;
    if (child.nodeType === 1 && child.localName === localName) return child;
  }
  return undefined;
}

function drawingElement(doc: Document, localName: string): Element {
  return doc.createElementNS(WP_NS, `wp:${localName}`);
}

function positionElement(doc: Document, horizontal: boolean, align: string): Element {
  const position = drawingElement(doc, horizontal ? 'positionH' : 'positionV');
  position.setAttribute('relativeFrom', horizontal ? 'margin' : 'paragraph');
  const alignEl = drawingElement(doc, 'align');
  alignEl.appendChild(doc.createTextNode(align));
  position.appendChild(alignEl);
  return position;
}

function buildFloatingAnchor(doc: Document, floating: ChartFloatingStyle): Element {
  const horizontal = floating.horizontalAlign ?? 'center';
  const vertical = floating.verticalAlign === 'middle' ? 'center' : floating.verticalAlign ?? 'top';
  const anchor = drawingElement(doc, 'anchor');
  anchor.setAttribute('distT', '0');
  anchor.setAttribute('distB', '0');
  anchor.setAttribute('distL', '114300');
  anchor.setAttribute('distR', '114300');
  anchor.setAttribute('simplePos', '0');
  anchor.setAttribute('relativeHeight', '251658240');
  anchor.setAttribute('behindDoc', floating.behindText ? '1' : '0');
  anchor.setAttribute('locked', '0');
  anchor.setAttribute('layoutInCell', '1');
  anchor.setAttribute('allowOverlap', '0');
  const simple = drawingElement(doc, 'simplePos');
  simple.setAttribute('x', '0');
  simple.setAttribute('y', '0');
  anchor.appendChild(simple);
  anchor.appendChild(positionElement(doc, true, horizontal));
  anchor.appendChild(positionElement(doc, false, vertical));
  return anchor;
}

const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function chartRelationshipIds(drawing: Element): string[] {
  const charts: Element[] = [];
  const scoped = drawing.getElementsByTagNameNS(CHART_NS, 'chart');
  for (let i = 0; i < scoped.length; i++) charts.push(scoped[i] as Element);
  const prefixed = drawing.getElementsByTagName('c:chart');
  for (let i = 0; i < prefixed.length; i++) charts.push(prefixed[i] as Element);
  const ids: string[] = [];
  for (const chart of charts) {
    const id = chart.getAttribute('r:id') || chart.getAttributeNS(REL_NS, 'id') || undefined;
    if (id && !ids.includes(id)) ids.push(id);
  }
  return ids;
}

function chartDrawingFor(nodeXml: Element, chartPath: string, chartRId?: string): Element {
  const drawings = nodeXml.getElementsByTagNameNS(WML_NS, 'drawing');
  const candidates: Element[] = [];
  for (let i = 0; i < drawings.length; i++) {
    const drawing = drawings[i] as Element;
    const ids = chartRelationshipIds(drawing);
    if (ids.length > 0 || drawing.getElementsByTagNameNS(CHART_NS, 'chart').length > 0 || drawing.getElementsByTagName('c:chart').length > 0) {
      candidates.push(drawing);
    }
  }
  const prefixed = nodeXml.getElementsByTagName('w:drawing');
  if (candidates.length === 0 && prefixed.length > 0) {
    for (let i = 0; i < prefixed.length; i++) {
      const drawing = prefixed[i] as Element;
      if (drawing.getElementsByTagName('c:chart').length > 0) candidates.push(drawing);
    }
  }
  let drawing = chartRId !== undefined
    ? candidates.find((candidate) => chartRelationshipIds(candidate).includes(chartRId))
    : candidates[0];
  if (!drawing && candidates.length === 1) drawing = candidates[0];
  if (!drawing) {
    throw styleRefusal(
      { chartPath, chartType: 'chart', seriesCount: 0, is3D: false },
      'chart node has no drawing that references the chart part',
      'Re-read the chart via inspect_document; if the chart drawing is gone, delete and re-insert the chart.'
    );
  }
  return drawing;
}

/** Convert or retarget the chart drawing wrapper without touching the chart part. */
export function applyChartFloating(nodeXml: Element, floating: ChartFloatingStyle, chartPath: string, chartRId?: string): void {
  const doc = nodeXml.ownerDocument;
  const drawing = chartDrawingFor(nodeXml, chartPath, chartRId);
  const inline = drawing.getElementsByTagNameNS(WP_NS, 'inline')[0] ?? drawing.getElementsByTagName('wp:inline')[0];
  const existingAnchor = drawing.getElementsByTagNameNS(WP_NS, 'anchor')[0] ?? drawing.getElementsByTagName('wp:anchor')[0];
  if (inline) {
    const anchor = buildFloatingAnchor(doc, floating);
    const extent = directChildByLocal(inline, 'extent');
    const effect = directChildByLocal(inline, 'effectExtent');
    const docPr = directChildByLocal(inline, 'docPr');
    const graphicFramePr = directChildByLocal(inline, 'cNvGraphicFramePr');
    const graphic = directChildByLocal(inline, 'graphic');
    const extLst = directChildByLocal(inline, 'extLst');
    if (extent) anchor.appendChild(extent);
    if (effect) anchor.appendChild(effect);
    const wrap = drawingElement(doc, floating.wrap === 'none' ? 'wrapNone' : 'wrapSquare');
    if (floating.wrap !== 'none') wrap.setAttribute('wrapText', 'bothSides');
    anchor.appendChild(wrap);
    if (docPr) anchor.appendChild(docPr);
    if (graphicFramePr) anchor.appendChild(graphicFramePr);
    if (graphic) anchor.appendChild(graphic);
    if (extLst) anchor.appendChild(extLst);
    drawing.replaceChild(anchor, inline);
    return;
  }
  if (!existingAnchor) {
    throw styleRefusal({ chartPath, chartType: 'chart', seriesCount: 0, is3D: false }, 'chart drawing is neither inline nor anchored', 'Re-read the chart via inspect_document; if the drawing is malformed, delete and re-insert the chart.');
  }
  if (floating.behindText !== undefined) existingAnchor.setAttribute('behindDoc', floating.behindText ? '1' : '0');
  if (floating.horizontalAlign !== undefined) {
    replacePosition(doc, existingAnchor, true, floating.horizontalAlign);
  }
  if (floating.verticalAlign !== undefined) {
    replacePosition(doc, existingAnchor, false, floating.verticalAlign === 'middle' ? 'center' : floating.verticalAlign);
  }
  if (floating.wrap !== undefined) {
    const oldWrap = directChildByLocal(existingAnchor, 'wrapNone') ?? directChildByLocal(existingAnchor, 'wrapSquare')
      ?? directChildByLocal(existingAnchor, 'wrapTight') ?? directChildByLocal(existingAnchor, 'wrapThrough')
      ?? directChildByLocal(existingAnchor, 'wrapTopAndBottom');
    const wrap = drawingElement(doc, floating.wrap === 'none' ? 'wrapNone' : 'wrapSquare');
    if (floating.wrap !== 'none') wrap.setAttribute('wrapText', 'bothSides');
    if (oldWrap) existingAnchor.replaceChild(wrap, oldWrap);
    else {
      const docPr = directChildByLocal(existingAnchor, 'docPr');
      if (docPr) existingAnchor.insertBefore(wrap, docPr);
      else existingAnchor.appendChild(wrap);
    }
  }
}

function replacePosition(doc: Document, anchor: Element, horizontal: boolean, align: string): void {
  const local = horizontal ? 'positionH' : 'positionV';
  const old = directChildByLocal(anchor, local);
  const fresh = positionElement(doc, horizontal, align);
  if (old) anchor.replaceChild(fresh, old);
  else anchor.appendChild(fresh);
}

/** Reads anchor placement from a chart's paragraph/drawing element. */
export function readChartFloating(nodeXml: Element): ChartFloatingStyle | undefined {
  const anchor = nodeXml.getElementsByTagNameNS(WP_NS, 'anchor')[0] ?? nodeXml.getElementsByTagName('wp:anchor')[0];
  if (!anchor) return undefined;
  const out: ChartFloatingStyle = {};
  const behind = anchor.getAttribute('behindDoc');
  if (behind !== null && behind !== undefined) out.behindText = behind === '1' || behind === 'true';
  const h = directChildByLocal(anchor, 'positionH');
  const v = directChildByLocal(anchor, 'positionV');
  const alignText = (position: Element | undefined): string | undefined => {
    if (!position) return undefined;
    const align = directChildByLocal(position, 'align');
    const text = align?.textContent?.trim();
    return text || undefined;
  };
  const ha = alignText(h);
  if (ha === 'left' || ha === 'center' || ha === 'right') out.horizontalAlign = ha;
  const va = alignText(v);
  if (va === 'top' || va === 'center' || va === 'bottom') out.verticalAlign = va === 'center' ? 'middle' : va;
  if (directChildByLocal(anchor, 'wrapNone')) out.wrap = 'none';
  else if (directChildByLocal(anchor, 'wrapSquare')) out.wrap = 'square';
  return Object.keys(out).length > 0 ? out : undefined;
}
