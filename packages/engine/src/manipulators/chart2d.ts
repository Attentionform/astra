import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST } from '../ast';
import { DocxParser, DocxZipAdapter } from '../parser';
import { RelationshipsManipulator } from './relationships';
import { createRelationshipsElement, createContentTypesElement, allocateNextDrawingDocPrId } from '../ooxml/ns';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { chartElementDescendants, chartPointTextsByIndex, chartScatterXValuesBySeries, chartTitleElement, detectBar2DChartTypeFromXml, detectChartKindFromXml, pairScatterPointTexts, scatterSeriesXList, scatterXOnlyValues } from '../parser/factory';
import { readChartStyle, type ChartFloatingStyle } from '../operations/chart-style';
import { generateEmbeddedWorkbook } from './chart3d';
import { buildChartDrawingContainer } from './chart-drawing';
import { DEFAULT_SERIES_PALETTE, defaultSeriesColor } from './chart3d';
// Series colors default per-series from the shared palette (chart3d.ts) —
// explicit series colors always win. The same palette drives pie fills.
import type { ChartXmlSummary } from './chart3d';
import { type Chart2DBarType, barChartVariantOf } from './chart2d-types';
export { BAR_2D_VARIANTS, barChartVariantOf } from './chart2d-types';
export type { BarChartVariant, Chart2DBarType } from './chart2d-types';

/**
 * Requestable 2D chart types. Base families keep their short names; stacked
 * variants append the OOXML grouping to the family (columnStacked writes
 * <c:grouping val="stacked"/>, columnPercentStacked writes
 * <c:grouping val="percentStacked"/>), matching the column3D/bar3D suffix
 * convention. All variants of one family share c:barChart and differ only in
 * c:barDir/c:grouping. Doughnut is the pie family with a hole: same series and
 * percent labels inside <c:doughnutChart> plus <c:holeSize val="50"/>.
 * Scatter uses numeric x categories plus per-series y values inside
 * <c:scatterChart>, with a value/value (c:valAx) axis pair.
 */
export type Chart2DType = Chart2DBarType | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'combo';
export const SUPPORTED_CHART_2D_TYPES = [
  'column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'combo',
  'columnStacked', 'barStacked', 'columnPercentStacked', 'barPercentStacked',
] as const;

export interface Chart2DSeries {
  name: string;
  values: number[];
  color?: string; // HEX color without '#' e.g. '0066CC'
  /** Scatter only: per-series numeric x values. Falls back to categories. */
  xValues?: number[];
  /** Combo only: plot family for this series. Defaults to 'column'. */
  chartType?: 'column' | 'line';
  /** Combo only: axis group. Secondary is supported for line series. */
  axis?: 'primary' | 'secondary';
}

export interface CreateChart2DOptions {
  chartType?: Chart2DType;
  title?: string;
  categories?: string[];
  series?: Chart2DSeries[];
  /** Doughnut hole diameter percentage (integer 1..90; default 50). */
  holeSize?: number;
  /** Internal: transform generated chart XML before it is written. */
  transformChartXml?: (xml: string) => string;
  /** Floating placement; omitted keeps the legacy inline drawing. */
  floating?: ChartFloatingStyle;
}

export interface InsertChart2DOptions extends CreateChart2DOptions {
  anchorNodeId?: string;
  id?: string;
  widthEmu?: number; // Default 5486400 (6 inches)
  heightEmu?: number; // Default 2800000 (3.06 inches)
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

function escapeXml(str: string): string {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Numeric values as a referenced cache (c:numCache) so renderers that
 * only read the cache (Word/LO) show data. The formula points at the embedded
 * workbook sheet written by insert2DChart (Sheet1, column B+ per series).
 */
function buildNumCacheXml(values: number[], seriesIdx: number): string {
  const col = chartColLetter(seriesIdx + 1);
  const lastRow = 1 + Math.max(values.length, 1);
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return `<c:numRef><c:f>Sheet1!$${col}$2:$${col}$${lastRow}</c:f>` +
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef>`;
}

const NUMERIC_TEXT = /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/;

function scatterDataError(message: string): Error {
  return Object.assign(new Error(message), {
    code: 'E_VALIDATION',
    retryable: false,
    remedy: 'Pass categories as numeric x values, one per point (for example ["0","2","4"]), or give every series its own xValues with one x per y value; a series with xValues uses it instead of categories.',
  });
}

/** Parses a numeric scatter list; every entry must be finite (1e999 is not). */
function parseScatterNumbers(
  values: readonly (string | number)[] | undefined,
  label: string,
  required: boolean
): number[] {
  if (!Array.isArray(values) || (required && values.length === 0)) {
    throw scatterDataError(
      required
        ? 'scatter chart x values are required: categories must hold at least one non-empty numeric x value per point.'
        : `${label} must contain one numeric x value per point.`
    );
  }
  const out: number[] = [];
  for (let i = 0; i < values.length; i++) {
    const rawValue = values[i];
    const raw = typeof rawValue === 'number' ? String(rawValue) : typeof rawValue === 'string' ? rawValue.trim() : '';
    if (raw.length === 0 || !NUMERIC_TEXT.test(raw)) {
      throw scatterDataError(`scatter chart ${label} at index ${i} ("${raw}") is not numeric.`);
    }
    const value = Number(raw);
    if (!Number.isFinite(value)) {
      throw scatterDataError(`scatter chart ${label} at index ${i} ("${raw}") is not finite.`);
    }
    out.push(value);
  }
  return out;
}

/**
 * Validates scatter data and returns each series' own numeric x list. A series
 * without `xValues` uses the shared `categories` list; a series with
 * `xValues` (a legal per-series c:xVal in OOXML) keeps it, so series may have
 * different point counts. Every series' y list must match its own x list.
 */
export function assertScatterData(
  categories: readonly string[] | undefined,
  series: readonly Chart2DSeries[] | undefined
): number[][] {
  if (!Array.isArray(series) || series.length === 0) {
    throw scatterDataError('scatter charts need at least one series with numeric y values.');
  }
  // categories are the shared fallback; when every series supplies its own
  // xValues a scatter payload is complete without a categories list.
  const needsSharedX = series.some((ser) => ser?.xValues === undefined);
  const sharedXValues = needsSharedX
    ? parseScatterNumbers(categories, 'x value', true)
    : [];
  const perSeriesX: number[][] = [];
  series.forEach((ser) => {
    const name = ser?.name ?? 'Series';
    const xValues = ser?.xValues !== undefined
      ? parseScatterNumbers(ser.xValues, `series "${name}" x value`, false)
      : sharedXValues;
    if (!Array.isArray(ser?.values) || ser.values.length !== xValues.length) {
      throw scatterDataError(
        `scatter chart series "${name}" has ${Array.isArray(ser?.values) ? ser.values.length : 0} y values but ${xValues.length} x values; x and y lists must be the same length.`
      );
    }
    if (ser.values.some((value: number) => typeof value !== 'number' || !Number.isFinite(value))) {
      throw scatterDataError(`scatter chart series "${name}" has a non-finite y value.`);
    }
    perSeriesX.push(xValues);
  });
  return perSeriesX;
}

/** Inline numeric list for scatter x/y values (no workbook reference). */
function buildNumLitXml(values: number[]): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numLit>`;
}

function buildSolidFillSpPr(color: string | undefined): string {
  const fill = (color || '0066CC').replace(/^#/, '');
  return `<c:spPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></c:spPr>`;
}

/** Visible line stroke for line series (default 1pt = 25400 EMU). */
function buildLineSpPr(color: string | undefined, widthEmu = 25400): string {
  const stroke = (color || '0066CC').replace(/^#/, '');
  return `<c:spPr><a:ln w="${widthEmu}"><a:solidFill><a:srgbClr val="${stroke}"/></a:solidFill></a:ln></c:spPr>`;
}

/** Per-series circle marker tinted with the series color. */
function buildMarkerXml(color: string | undefined): string {
  const fill = (color || '0066CC').replace(/^#/, '');
  return `<c:marker><c:symbol val="circle"/><c:spPr><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill></c:spPr></c:marker>`;
}


function comboDataError(message: string, remedy: string): Error {
  return Object.assign(new Error(message), { code: 'E_VALIDATION', retryable: false, remedy });
}

function buildComboCatXml(categories: string[]): string {
  const pts = categories.map((cat, i) => `<c:pt idx="${i}"><c:v>${escapeXml(cat)}</c:v></c:pt>`).join('');
  return `<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${pts}</c:strLit></c:cat>`;
}

function comboPlotSeriesXml(
  ser: Chart2DSeries,
  globalIndex: number,
  kind: 'column' | 'line',
  categories: string[]
): string {
  const color = defaultSeriesColor(ser.color, globalIndex);
  const shape = kind === 'line' ? buildLineSpPr(color) + buildMarkerXml(color) : buildSolidFillSpPr(color);
  return `<c:ser>
      <c:idx val="${globalIndex}"/>
      <c:order val="${globalIndex}"/>
      <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
      ${shape}
      ${buildComboCatXml(categories)}
      <c:val>${buildNumCacheXml(ser.values, globalIndex)}</c:val>
    </c:ser>`;
}

function comboAxisXml(kind: 'cat' | 'val', id: number, cross: number, opts: {
  pos: string;
  hidden?: boolean;
  gridlines?: boolean;
  crosses?: 'autoZero' | 'max';
}): string {
  if (kind === 'cat') {
    return `<c:catAx>
      <c:axId val="${id}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="${opts.hidden ? '1' : '0'}"/>
      <c:axPos val="${opts.pos}"/>
      <c:majorTickMark val="out"/>
      <c:minorTickMark val="none"/>
      <c:tickLblPos val="nextTo"/>
      <c:crossAx val="${cross}"/>
      <c:auto val="1"/>
      <c:lblAlgn val="ctr"/>
      <c:lblOffset val="100"/>
      <c:noMultiLvlLbl val="0"/>
    </c:catAx>`;
  }
  return `<c:valAx>
    <c:axId val="${id}"/>
    <c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="${opts.hidden ? '1' : '0'}"/>
    <c:axPos val="${opts.pos}"/>
    ${opts.gridlines ? '<c:majorGridlines/>' : ''}
    <c:numFmt formatCode="General" sourceLinked="1"/>
    <c:majorTickMark val="out"/>
    <c:minorTickMark val="none"/>
    <c:tickLblPos val="nextTo"/>
    <c:crossAx val="${cross}"/>
    <c:crosses val="${opts.crosses ?? 'autoZero'}"/>
    <c:crossBetween val="between"/>
  </c:valAx>`;
}


function normalizeComboSeries(series: readonly Chart2DSeries[] | undefined): Array<{
  ser: Chart2DSeries;
  index: number;
  kind: 'column' | 'line';
  axis: 'primary' | 'secondary';
}> {
  const input = series && series.length > 0 ? series : [];
  if (input.length < 2) {
    throw comboDataError(
      'combo charts need at least two series (one column and one line).',
      'Pass series with chartType "column" and "line", e.g. [{name:"Revenue",values:[1,2],chartType:"column"},{name:"Margin",values:[.2,.3],chartType:"line"}].'
    );
  }
  const normalized = input.map((ser, index) => {
    const kind = ser.chartType ?? 'column';
    const axis = ser.axis ?? 'primary';
    if (kind !== 'column' && kind !== 'line') {
      throw comboDataError(
        `combo series ${index + 1} has unsupported chartType "${String(kind)}".`,
        'Use chartType "column" or "line" on every combo series.'
      );
    }
    if (axis !== 'primary' && axis !== 'secondary') {
      throw comboDataError(
        `combo series ${index + 1} has unsupported axis "${String(axis)}".`,
        'Use axis "primary" or "secondary" on every combo series.'
      );
    }
    if (axis === 'secondary' && kind !== 'line') {
      throw comboDataError(
        `combo series ${index + 1} is a ${kind} on the secondary axis.`,
        'Secondary axes are supported for line series only.'
      );
    }
    return { ser, index, kind, axis } as const;
  });
  if (!normalized.some((entry) => entry.kind === 'column')) {
    throw comboDataError('combo charts need at least one column series.', 'Set chartType:"column" on at least one series.');
  }
  if (!normalized.some((entry) => entry.kind === 'line')) {
    throw comboDataError('combo charts need at least one line series.', 'Set chartType:"line" on at least one series.');
  }
  return normalized;
}

export function assertComboData(series: readonly Chart2DSeries[] | undefined): void {
  normalizeComboSeries(series);
}

/** Generates the supported combo part: column(s) + line(s), optional secondary line axis. */
function generateCombo2DChartXml(
  options: CreateChart2DOptions,
  titleXml: string,
  externalDataRId?: string
): string {
  const categories = options.categories && options.categories.length > 0
    ? options.categories
    : ['Q1', 'Q2', 'Q3', 'Q4'];
  const normalized = normalizeComboSeries(options.series);
  const primaryBars = normalized.filter((entry) => entry.kind === 'column' && entry.axis === 'primary');
  const primaryLines = normalized.filter((entry) => entry.kind === 'line' && entry.axis === 'primary');
  const secondaryLines = normalized.filter((entry) => entry.kind === 'line' && entry.axis === 'secondary');
  const hasSecondary = secondaryLines.length > 0;

  const barPlot = primaryBars.length === 0 ? '' : `<c:barChart>
    <c:barDir val="col"/>
    <c:grouping val="clustered"/>
    <c:varyColors val="0"/>
    ${primaryBars.map((entry) => comboPlotSeriesXml(entry.ser, entry.index, 'column', categories)).join('')}
    <c:gapWidth val="150"/>
    <c:overlap val="-30"/>
    <c:axId val="1001"/>
    <c:axId val="1002"/>
  </c:barChart>`;
  const linePlot = (series: typeof normalized, primary: boolean): string => series.length === 0 ? '' : `<c:lineChart>
    <c:grouping val="standard"/>
    <c:varyColors val="0"/>
    ${series.map((entry) => comboPlotSeriesXml(entry.ser, entry.index, 'line', categories)).join('')}
    <c:marker val="1"/>
    <c:smooth val="0"/>
    <c:axId val="${primary ? '1001' : '2001'}"/>
    <c:axId val="${primary ? '1002' : '2002'}"/>
  </c:lineChart>`;
  const plots = `${barPlot}${linePlot(primaryLines, true)}${linePlot(secondaryLines, false)}`;
  const primaryAxes = `${comboAxisXml('cat', 1001, 1002, { pos: 'b' })}${comboAxisXml('val', 1002, 1001, { pos: 'l', gridlines: true })}`;
  const secondaryAxes = hasSecondary
    ? `${comboAxisXml('cat', 2001, 2002, { pos: 'b', hidden: true })}${comboAxisXml('val', 2002, 2001, { pos: 'r', crosses: 'max' })}`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:chart>${titleXml}
    <c:plotArea>
      <c:layout/>
      ${plots}
      ${primaryAxes}${secondaryAxes}
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
  ${externalDataRId ? `<c:externalData r:id="${externalDataRId}"><c:autoUpdate val="0"/></c:externalData>` : ''}
</c:chartSpace>`;
}

export class Chart2DManipulator {
  private ast: DocumentAST;
  private parser?: DocxParser;
  private domParser: DOMParser;
  private serializer: XMLSerializer;

  constructor(ast: DocumentAST, parser?: DocxParser) {
    this.ast = ast;
    this.parser = parser;
    this.domParser = new DOMParser();
    this.serializer = new XMLSerializer();
  }

  /**
   * Generates ISO/IEC 29500-1 Clause 21.2 compliant <c:chartSpace> XML string for 2D charts.
   */
  public generate2DChartXml(options: CreateChart2DOptions, externalDataRId?: string): string {
    // An absent title keeps the historical default; an empty title is an
    // explicit titleless part and must stay titleless through a structural
    // regeneration (no <c:title> element at all).
    const title = options.title === undefined ? 'Chart Analysis' : options.title;
    const titleXml = title.length > 0
      ? `
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr b="1" sz="1400"/>
              <a:t>${escapeXml(title)}</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>`
      : '';

    const rawChartType = options.chartType || 'column';
    const supportedChartTypes = SUPPORTED_CHART_2D_TYPES as readonly string[];
    if (!supportedChartTypes.includes(rawChartType)) {
      throw Object.assign(
        new Error(`Unsupported chart type "${rawChartType}". Supported types: ${SUPPORTED_CHART_2D_TYPES.join(', ')}`),
        {
          code: 'E_VALIDATION',
          retryable: false,
          remedy: `Use one of the supported chart types: ${SUPPORTED_CHART_2D_TYPES.join(', ')}`,
        }
      );
    }
    const chartType = rawChartType as Chart2DType;
    if (options.holeSize !== undefined) {
      if (chartType !== 'doughnut') {
        throw Object.assign(
          new Error('holeSize is only supported for doughnut charts.'),
          { code: 'E_VALIDATION', retryable: false, remedy: 'Omit holeSize unless chartType is doughnut.' }
        );
      }
      if (!Number.isInteger(options.holeSize) || options.holeSize < 1 || options.holeSize > 90) {
        throw Object.assign(
          new Error('holeSize must be an integer from 1 to 90.'),
          { code: 'E_VALIDATION', retryable: false, remedy: 'Use an integer holeSize from 1 to 90 (50 = half-diameter hole), or omit it for the default 50.' }
        );
      }
    }
    const holeSize = options.holeSize ?? 50;
    const scatterSeriesXValues = chartType === 'scatter'
      ? assertScatterData(options.categories, options.series)
      : undefined;

    const categories = options.categories && options.categories.length > 0
      ? options.categories
      : ['Q1', 'Q2', 'Q3', 'Q4'];

    const seriesList: Chart2DSeries[] = options.series && options.series.length > 0
      ? options.series
      : [{ name: 'Performance', values: [12.5, 14.8, 18.2, 22.0], color: '0066CC' }];

    const buildCatXml = (): string => {
      let catPts = '';
      categories.forEach((cat, cIdx) => {
        catPts += `<c:pt idx="${cIdx}"><c:v>${escapeXml(cat)}</c:v></c:pt>`;
      });
      return `<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat>`;
    };

    if (chartType === 'combo') {
      return generateCombo2DChartXml(options, titleXml, externalDataRId);
    }

    const barVariant = barChartVariantOf(chartType);
    const isPieFamily = chartType === 'pie' || chartType === 'doughnut';
    let chartTypeXml = '';

    if (isPieFamily) {
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        // Per-point fills so slices actually differ (varyColors=1 alone
        // renders every slice the series color in several renderers).
        const baseFill = defaultSeriesColor(ser.color, sIdx);
        let dPts = '';
        ser.values.forEach((_v, vIdx) => {
          const fill = vIdx === 0 ? baseFill : DEFAULT_SERIES_PALETTE[vIdx % DEFAULT_SERIES_PALETTE.length];
          dPts += `<c:dPt><c:idx val="${vIdx}"/>${buildSolidFillSpPr(fill)}</c:dPt>`;
        });

        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          ${buildSolidFillSpPr(baseFill)}
          ${dPts}
          ${buildCatXml()}
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
        </c:ser>`;
      });

      // Doughnut is the same pie family with a hole: c:holeSize is the hole
      // diameter as a percentage of the chart diameter (default 50 = half).
      const plotTag = chartType === 'doughnut' ? 'doughnutChart' : 'pieChart';
      const holeSizeXml = chartType === 'doughnut' ? `\n        <c:holeSize val="${holeSize}"/>` : '';
      chartTypeXml = `<c:${plotTag}>
        <c:varyColors val="1"/>
        ${seriesXml}
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="1"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:firstSliceAng val="0"/>${holeSizeXml}
      </c:${plotTag}>`;
    } else if (chartType === 'line') {
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          ${buildLineSpPr(defaultSeriesColor(ser.color, sIdx))}
          ${buildMarkerXml(defaultSeriesColor(ser.color, sIdx))}
          ${buildCatXml()}
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
          <c:smooth val="0"/>
        </c:ser>`;
      });

      chartTypeXml = `<c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        ${seriesXml}
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:lineChart>`;
    } else if (chartType === 'area') {
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          ${buildSolidFillSpPr(defaultSeriesColor(ser.color, sIdx))}
          ${buildCatXml()}
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
        </c:ser>`;
      });

      chartTypeXml = `<c:areaChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        ${seriesXml}
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:areaChart>`;
    } else if (chartType === 'scatter') {
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        const color = defaultSeriesColor(ser.color, sIdx);
        const xValues = scatterSeriesXValues?.[sIdx] ?? [];
        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          ${buildLineSpPr(color)}
          ${buildMarkerXml(color)}
          <c:xVal>${buildNumLitXml(xValues)}</c:xVal>
          <c:yVal>${buildNumLitXml(ser.values)}</c:yVal>
          <c:smooth val="0"/>
        </c:ser>`;
      });

      chartTypeXml = `<c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:varyColors val="0"/>
        ${seriesXml}
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:scatterChart>`;
    } else if (barVariant) {
      // column (vertical) and bar (horizontal) share c:barChart, differing in
      // barDir; stacked/percentStacked differ only in c:grouping. Stacked bars
      // must overlap fully or renderers draw them side by side.
      let seriesXml = '';
      seriesList.forEach((ser, sIdx) => {
        const fillColor = defaultSeriesColor(ser.color, sIdx);

        seriesXml += `<c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${escapeXml(ser.name)}</c:v></c:tx>
          <c:spPr>
            <a:solidFill>
              <a:srgbClr val="${fillColor}"/>
            </a:solidFill>
          </c:spPr>
          ${buildCatXml()}
          <c:val>${buildNumCacheXml(ser.values, sIdx)}</c:val>
        </c:ser>`;
      });

      const barDir = barVariant.barDir;
      const gapWidth = barDir === 'bar' ? 120 : 150;
      const overlap = barVariant.grouping === 'clustered'
        ? (barDir === 'bar' ? undefined : -30)
        : 100;
      const spacingXml = `<c:gapWidth val="${gapWidth}"/>` +
        (overlap !== undefined ? `\n      <c:overlap val="${overlap}"/>` : '');
      chartTypeXml = `<c:barChart>
        <c:barDir val="${barDir}"/>
        <c:grouping val="${barVariant.grouping}"/>
        <c:varyColors val="0"/>
        ${seriesXml}
        ${spacingXml}
        <c:axId val="1001"/>
        <c:axId val="1002"/>
      </c:barChart>`;
    }

    // Horizontal bars transpose the axes: categories run vertically (left),
    // values run horizontally (bottom).
    const isHorizontalBar = barVariant?.barDir === 'bar';
    const catAxPos = isHorizontalBar ? 'l' : 'b';
    const valAxPos = isHorizontalBar ? 'b' : 'l';
    const isPercentStacked = barVariant?.percent === true;
    const isScatter = chartType === 'scatter';

    const categoryAxisXml = isPieFamily || isScatter ? '' : `<c:catAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${catAxPos}"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>`;

    const valueAxisXml = isPieFamily || isScatter ? '' : `<c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="${valAxPos}"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="${isPercentStacked ? '0%' : 'General'}" sourceLinked="${isPercentStacked ? '0' : '1'}"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>`;

    // Scatter carries two value axes: x at the bottom, y at the left, each
    // pointing at the other's id. Unlike catAx, both scale numerically.
    const scatterAxesXml = isScatter ? `<c:valAx>
        <c:axId val="1001"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1002"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="1002"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="1001"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>` : '';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:chart>${titleXml}
    <c:plotArea>
      <c:layout/>
      ${chartTypeXml}
      ${categoryAxisXml}${scatterAxesXml}
      ${valueAxisXml}
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
  </c:chart>
  ${externalDataRId ? `<c:externalData r:id="${externalDataRId}"><c:autoUpdate val="0"/></c:externalData>` : ''}
</c:chartSpace>`;
  }

  /**
   * Generates DrawingML inline wrapper for embedding chart in document.xml.
   * Carries the P1 centered, auto-spaced w:pPr so Word expands the line box
   * for the inline graphic even when Normal uses exact line spacing.
   */
  public generateDrawingEmbedXml(rId: string, docPrId: number = 1, name: string = 'Chart 1', widthEmu: number = 5486400, heightEmu: number = 2800000, floating?: ChartFloatingStyle): string {
    const safeName = escapeXml(name || 'Chart');

    return `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
     xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:pPr><w:spacing w:before="140" w:after="140" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>
  <w:r>
    <w:drawing>
      ${buildChartDrawingContainer({ rId, docPrId, name: safeName, widthEmu, heightEmu, floating })}
    </w:drawing>
  </w:r>
</w:p>`;
  }

  /**
   * Inserts a complete 2D chart into the document package and AST.
   */
  public insert2DChart(options: InsertChart2DOptions): { rId: string; chartPath: string; drawingNodeId?: string; workbookPath?: string } {
    const rawChartType = options.chartType || 'column';
    const supportedChartTypes = SUPPORTED_CHART_2D_TYPES as readonly string[];
    if (!supportedChartTypes.includes(rawChartType)) {
      throw Object.assign(
        new Error(`Unsupported chart type "${rawChartType}". Supported types: ${supportedChartTypes.join(', ')}`),
        {
          code: 'E_VALIDATION',
          retryable: false,
          remedy: `Use one of the supported chart types: ${supportedChartTypes.join(', ')}`,
        }
      );
    }
    // Validate before any package mutation (relationship, workbook, part).
    const scatterSeriesXValues = rawChartType === 'scatter'
      ? assertScatterData(options.categories, options.series)
      : undefined;
    if (rawChartType === 'combo') assertComboData(options.series);

    const relsManipulator = new RelationshipsManipulator(this.ast);

    // Calculate next chart part number
    let maxNum = 1;
    for (const rel of Object.values(this.ast.relationships)) {
      if (rel.type.includes('chart')) {
        const match = rel.target.match(/chart(\d+)\.xml/);
        if (match) {
          const num = parseInt(match[1], 10);
          if (num >= maxNum) maxNum = num + 1;
        }
      }
    }

    const chartPartName = `charts/chart${maxNum}.xml`;
    const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
    const rId = relsManipulator.addRelationship(relType, chartPartName);

    const categories = options.categories && options.categories.length > 0
      ? options.categories
      : ['Q1', 'Q2', 'Q3', 'Q4'];

    const seriesList: Chart2DSeries[] = options.series && options.series.length > 0
      ? options.series
      : [{ name: 'Performance', values: [12.5, 14.8, 18.2, 22.0], color: '0066CC' }];

    const widthEmu = options.widthEmu || 5486400;
    const heightEmu = options.heightEmu || 2800000;

    // Globally unique docPr id via shared DOM-scanning allocator
    // (images + charts occupy the same id space; scans wp:docPr + pic:cNvPr).
    let parserDom: Document | null = null;
    try {
      parserDom = this.parser ? this.parser.getDocumentDom() : null;
    } catch { parserDom = null; }
    const docPrId = allocateNextDrawingDocPrId(parserDom);

    const drawingXml = this.generateDrawingEmbedXml(rId, docPrId, options.title || `Chart ${maxNum}`, widthEmu, heightEmu, options.floating);

    if (this.parser) {
      const workbookPartName = `embeddings/Microsoft_Excel_Worksheet${maxNum}.xlsx`;
      const workbookBuf = generateEmbeddedWorkbook(seriesList, categories, scatterSeriesXValues);
      const zipAdapter = this.parser.getZipAdapter();

      // 1. Embed real Excel workbook binary
      zipAdapter.writeBuffer(`word/${workbookPartName}`, workbookBuf);

      // 2. Add chart relationship pointing to embedded Excel workbook
      const chartRelsPath = `word/charts/_rels/chart${maxNum}.xml.rels`;
      const chartRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${maxNum}.xlsx"/>
</Relationships>`;
      zipAdapter.writeText(chartRelsPath, chartRelsXml);

      // 3. Generate chart.xml with c:externalData referencing rId1
      const generatedXml = this.generate2DChartXml(options, 'rId1');
      const chartXmlStr = options.transformChartXml ? options.transformChartXml(generatedXml) : generatedXml;
      zipAdapter.writeText(`word/${chartPartName}`, chartXmlStr);

      // 4. Ensure relationship in word/_rels/document.xml.rels
      this.ensureDocumentRelationship(zipAdapter, rId, relType, chartPartName);

      // 5. Ensure content type overrides in [Content_Types].xml
      const chartContentType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
      this.ensureContentTypeOverride(zipAdapter, `/word/${chartPartName}`, chartContentType);

      const xlsxContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      this.ensureContentTypeOverride(zipAdapter, `/word/${workbookPartName}`, xlsxContentType);
      this.ensureContentTypeDefault(zipAdapter, 'xlsx', xlsxContentType);

      let drawingNodeId: string | undefined;
      if (options.anchorNodeId) {
        drawingNodeId = this.parser.insertRootNode(drawingXml, options.anchorNodeId, 'after', options.id);
      } else {
        drawingNodeId = this.parser.insertRootNode(drawingXml, undefined, 'append', options.id);
      }
      return { rId, chartPath: `word/${chartPartName}`, drawingNodeId, workbookPath: `word/${workbookPartName}` };
    }

    return { rId, chartPath: `word/${chartPartName}` };
  }

  private ensureDocumentRelationship(zipAdapter: DocxZipAdapter, rId: string, type: string, target: string) {
    const relsPath = 'word/_rels/document.xml.rels';
    let relsXml = zipAdapter.readTextOrNull(relsPath);
    if (relsXml) assertNoXxe(relsXml, relsPath);

    if (!relsXml) {
      relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${rId}" Type="${type}" Target="${target}"/>
</Relationships>`;
    } else {
      const doc = this.domParser.parseFromString(relsXml, 'text/xml');
      const relsRoot = doc.documentElement;
      let found = false;

      const existingRels = Array.from(relsRoot.getElementsByTagName('Relationship'));
      for (const rel of existingRels) {
        if (rel.getAttribute('Id') === rId) {
          rel.setAttribute('Type', type);
          rel.setAttribute('Target', target);
          found = true;
          break;
        }
      }

      if (!found) {
        const newRel = createRelationshipsElement(doc, 'Relationship');
        newRel.setAttribute('Id', rId);
        newRel.setAttribute('Type', type);
        newRel.setAttribute('Target', target);
        relsRoot.appendChild(newRel);
      }
      relsXml = this.serializer.serializeToString(doc);
    }

    zipAdapter.writeText(relsPath, relsXml);
  }

  private ensureContentTypeOverride(zipAdapter: DocxZipAdapter, partName: string, contentType: string) {
    const typesPath = '[Content_Types].xml';
    let typesXml = zipAdapter.readTextOrNull(typesPath);
    if (typesXml) assertNoXxe(typesXml, typesPath);

    if (!typesXml) {
      typesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="${partName}" ContentType="${contentType}"/>
</Types>`;
    } else {
      const doc = this.domParser.parseFromString(typesXml, 'text/xml');
      const typesRoot = doc.documentElement;
      let found = false;

      const overrides = Array.from(typesRoot.getElementsByTagName('Override'));
      for (const ov of overrides) {
        if (ov.getAttribute('PartName') === partName) {
          ov.setAttribute('ContentType', contentType);
          found = true;
          break;
        }
      }

      if (!found) {
        const newOv = createContentTypesElement(doc, 'Override');
        newOv.setAttribute('PartName', partName);
        newOv.setAttribute('ContentType', contentType);
        typesRoot.appendChild(newOv);
      }
      typesXml = this.serializer.serializeToString(doc);
    }

    zipAdapter.writeText(typesPath, typesXml);
  }

  private ensureContentTypeDefault(zipAdapter: DocxZipAdapter, extension: string, contentType: string) {
    const typesPath = '[Content_Types].xml';
    const typesXml = zipAdapter.readTextOrNull(typesPath);
    if (!typesXml) return;
    assertNoXxe(typesXml, typesPath);

    const doc = this.domParser.parseFromString(typesXml, 'text/xml');
    const typesRoot = doc.documentElement;
    const defaults = Array.from(typesRoot.getElementsByTagName('Default'));
    let found = false;
    for (const d of defaults) {
      if (d.getAttribute('Extension')?.toLowerCase() === extension.toLowerCase()) {
        d.setAttribute('ContentType', contentType);
        found = true;
        break;
      }
    }
    if (!found) {
      const newDef = createContentTypesElement(doc, 'Default');
      newDef.setAttribute('Extension', extension);
      newDef.setAttribute('ContentType', contentType);
      typesRoot.insertBefore(newDef, typesRoot.firstChild);
    }
    zipAdapter.writeText(typesPath, this.serializer.serializeToString(doc));
  }
}

/**
 * Read-only text helper for chart-part surfacing (no mutation).
 */
function chartXmlTextOf(el: Element | null | undefined): string {
  try {
    const t = el?.textContent;
    return typeof t === 'string' ? t : '';
  } catch {
    return '';
  }
}

/**
 * Parses an existing DrawingML chart part into a summary (read path only —
 * no write path in this file is touched or affected). Best-effort: malformed
 * XML yields an empty summary, never a throw.
 */
export function summarize2DChartXml(chartXml: string): ChartXmlSummary {
  const empty: ChartXmlSummary = { title: undefined, categories: [], series: [], chartType: undefined };
  if (typeof chartXml !== 'string' || chartXml.length === 0) return empty;
  try {
    const doc = new DOMParser().parseFromString(chartXml, 'text/xml');
    if (!doc || !doc.documentElement) return empty;
    const root = doc.documentElement;

    let title: string | undefined;
    try {
      const titleEl = chartTitleElement(root);
      if (titleEl) {
        const tEls = titleEl.getElementsByTagName('a:t');
        const parts: string[] = [];
        for (let i = 0; i < tEls.length; i++) parts.push(chartXmlTextOf(tEls[i] as Element));
        const joined = parts.join('').trim();
        title = joined ? joined : undefined;
      }
    } catch {
      title = undefined;
    }

    const categories: string[] = [];
    // Scatter semantics apply only when the chart resolves to the scatter
    // family. A legal combo part (e.g. bar+scatter) keeps the pre-existing
    // family's category read instead of adopting the scatter x lists.
    const isScatterSummary = detectChartKindFromXml(chartXml).chartType === 'scatter';
    const isComboSummary = detectChartKindFromXml(chartXml).chartType === 'combo';
    const scatterXText = isScatterSummary ? chartScatterXValuesBySeries(root) : [];
    try {
      if (isScatterSummary) {
        const firstX = scatterSeriesXList(scatterXText, 0);
        const catFallback = firstX === undefined
          ? (() => {
              const el = chartElementDescendants(root, 'cat')[0];
              const container = el
                ? chartElementDescendants(el, 'strCache')[0] ?? chartElementDescendants(el, 'strLit')[0]
                  ?? chartElementDescendants(el, 'numCache')[0] ?? chartElementDescendants(el, 'numLit')[0]
                : undefined;
              return container
                ? chartElementDescendants(container, 'pt').map((pt) => chartXmlTextOf(chartElementDescendants(pt, 'v')[0]))
                : [];
            })()
          : firstX;
        for (const value of catFallback) categories.push(value);
      } else {
        const catEl = chartElementDescendants(root, 'cat')[0];
        if (catEl) {
          const ptEls = chartElementDescendants(catEl, 'pt');
          for (let i = 0; i < ptEls.length; i++) {
            const vEls = chartElementDescendants(ptEls[i], 'v');
            if (vEls.length > 0) categories.push(chartXmlTextOf(vEls[0]));
          }
        }
      }
    } catch {
      /* categories stay as collected */
    }

    const series: Array<{ name: string; values: number[]; xValues?: number[]; unpairedX?: number; unpairedY?: number; chartType?: 'column' | 'line'; axis?: 'primary' | 'secondary' }> = [];
    try {
      const serEls = chartElementDescendants(root, 'ser');
      if (isComboSummary) {
        serEls.sort((a, b) => {
          const ai = Number(chartElementDescendants(a, 'idx')[0]?.getAttribute('val') ?? Number.NaN);
          const bi = Number(chartElementDescendants(b, 'idx')[0]?.getAttribute('val') ?? Number.NaN);
          return ai - bi;
        });
      }
      for (let s = 0; s < serEls.length; s++) {
        const ser = serEls[s];
        let name = 'Series';
        try {
          const txEls = chartElementDescendants(ser, 'tx');
          if (txEls.length > 0) {
            const vEls = chartElementDescendants(txEls[0], 'v');
            const nm = vEls.length > 0 ? chartXmlTextOf(vEls[0]).trim() : '';
            if (nm) name = nm;
          }
        } catch {
          /* name stays default */
        }
        const values: number[] = [];
        try {
          const valEl = chartElementDescendants(ser, 'val')[0] ??
            (isScatterSummary ? chartElementDescendants(ser, 'yVal')[0] : undefined);
          if (valEl) {
            const ptEls = chartElementDescendants(valEl, 'pt');
            for (let i = 0; i < ptEls.length; i++) {
              const vEls = chartElementDescendants(ptEls[i], 'v');
              if (vEls.length > 0) values.push(parseFloat(chartXmlTextOf(vEls[0])) || 0);
            }
          }
        } catch {
          /* values stay as collected */
        }
        const xOverride = isScatterSummary ? scatterSeriesXList(scatterXText, s) : undefined;
        let xValues = xOverride !== undefined ? xOverride.map((value) => parseFloat(value) || 0) : undefined;
        let unpairedX = 0;
        let unpairedY = 0;
        const pairValEl = isScatterSummary
          ? (chartElementDescendants(ser, 'val')[0] ?? chartElementDescendants(ser, 'yVal')[0])
          : undefined;
        if (isScatterSummary && xOverride !== undefined && pairValEl) {
          const paired = pairScatterPointTexts(xOverride, chartPointTextsByIndex(pairValEl));
          xValues = paired.xValues.length > 0 ? paired.xValues : undefined;
          values.length = 0;
          values.push(...paired.values);
          unpairedX = paired.unpairedX;
          unpairedY = paired.unpairedY;
        } else if (isScatterSummary && xOverride !== undefined) {
          // No y container: expose the x list and count every non-empty point
          // as unpaired, matching the parser and update readers.
          const xOnly = scatterXOnlyValues(xOverride);
          xValues = xOnly.xValues.length > 0 ? xOnly.xValues : undefined;
          unpairedX = xOnly.unpairedX;
        }
        let comboType: 'column' | 'line' | undefined;
        let comboAxis: 'primary' | 'secondary' | undefined;
        if (isComboSummary) {
          let plot: Element | null = ser.parentNode as Element | null;
          while (plot && plot.localName !== 'barChart' && plot.localName !== 'lineChart') plot = plot.parentNode as Element | null;
          if (plot) {
            comboType = plot.localName === 'lineChart' ? 'line' : 'column';
            const axIds = chartElementDescendants(plot, 'axId').map((el) => el.getAttribute('val'));
            comboAxis = axIds.includes('2001') || axIds.includes('2002') ? 'secondary' : 'primary';
          }
        }
        series.push({
          name,
          values,
          ...(xValues ? { xValues } : {}),
          ...(unpairedX > 0 ? { unpairedX } : {}),
          ...(unpairedY > 0 ? { unpairedY } : {}),
          ...(comboType ? { chartType: comboType } : {}),
          ...(comboAxis ? { axis: comboAxis } : {}),
        });
      }
    } catch {
      /* series stay as collected */
    }

    // Scatter categories mirror the first series' paired x list, so
    // categories.length always matches that series' xValues.length.
    if (isScatterSummary && series.length > 0 && series[0].xValues !== undefined) {
      categories.length = 0;
      for (const value of series[0].xValues) categories.push(String(value));
    }

    let chartType: string | undefined;
    try {
      const has = (localName: string): boolean => chartElementDescendants(root, localName).length > 0;
      if (has('doughnutChart')) chartType = 'doughnut';
      else if (has('pieChart')) chartType = 'pie';
      else if (has('barChart') && has('lineChart')) chartType = 'combo';
      else if (has('lineChart')) chartType = 'line';
      else if (has('areaChart')) chartType = 'area';
      else if (has('barChart')) {
        chartType = detectBar2DChartTypeFromXml(chartXml);
      }
      else if (has('scatterChart')) chartType = 'scatter';
    } catch {
      chartType = undefined;
    }
    const style = readChartStyle(chartXml);
    return { title, categories, series, chartType, ...(style !== undefined ? { style } : {}) };
  } catch {
    return empty;
  }
}

