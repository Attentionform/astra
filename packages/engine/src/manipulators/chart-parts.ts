/**
 * Shared XML fragments for chart parts. Every generator and patcher writes the
 * chart title and the major gridlines through these helpers so Word and
 * LibreOffice render the same chart.
 *
 * Why these exist (0.5.3):
 *  - A `<c:title>` without `<c:overlay val="0"/>` is drawn on top of the plot
 *    area by Word. The title block was written in four places and none of them
 *    set the overlay.
 *  - An empty `<c:majorGridlines/>` has no line style. Word draws it black,
 *    LibreOffice draws it light grey. A light grey line is written explicitly.
 */

/** Light grey used for gridlines when nothing else is requested. */
export const DEFAULT_GRIDLINE_COLOR = 'D9D9D9';

/** Gridline width in EMU (0.75 pt). */
const GRIDLINE_WIDTH_EMU = 9525;

export type ChartQualifier = (localName: string) => string;

const DEFAULT_QUALIFIER: ChartQualifier = (localName) => `c:${localName}`;

export function escapeChartXmlText(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export interface ChartTitleFormat {
  /** Size in hundredths of a point (1400 = 14 pt). */
  sizeHundredthsPt?: number;
  bold?: boolean;
  color?: string;
}

/**
 * The chart-level `<c:title>` with the text `title` (raw text, escaped here).
 * `<c:overlay val="0"/>` keeps the title above the plot area instead of on it.
 * The element order follows CT_Title: tx, layout, overlay, spPr, txPr.
 */
export function chartTitleXml(
  title: string,
  q: ChartQualifier = DEFAULT_QUALIFIER,
  format: ChartTitleFormat = {}
): string {
  return (
    `<${q('title')}>` +
    chartTitleTxXml(title, q, format) +
    `<${q('overlay')} val="0"/>` +
    `</${q('title')}>`
  );
}

/** Only the `<c:tx>` child of a chart title (rich text run), for in-place text replacement. */
export function chartTitleTxXml(
  title: string,
  q: ChartQualifier = DEFAULT_QUALIFIER,
  format: ChartTitleFormat = {}
): string {
  const size = format.sizeHundredthsPt ?? 1400;
  const bold = format.bold ?? true;
  const fill = format.color ? `<a:solidFill><a:srgbClr val="${format.color}"/></a:solidFill>` : '';
  return (
    `<${q('tx')}><${q('rich')}><a:bodyPr/><a:lstStyle/><a:p><a:r>` +
    `<a:rPr lang="en-US" b="${bold ? 1 : 0}" sz="${size}">${fill}</a:rPr>` +
    `<a:t>${escapeChartXmlText(title)}</a:t>` +
    `</a:r></a:p></${q('rich')}></${q('tx')}>`
  );
}

/**
 * Major gridlines with an explicit light grey 0.75 pt line, so the colour does
 * not depend on the renderer's default.
 */
export function majorGridlinesXml(
  q: ChartQualifier = DEFAULT_QUALIFIER,
  color: string = DEFAULT_GRIDLINE_COLOR
): string {
  return (
    `<${q('majorGridlines')}>` +
    `<${q('spPr')}><a:ln w="${GRIDLINE_WIDTH_EMU}" cap="flat"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></${q('spPr')}>` +
    `</${q('majorGridlines')}>`
  );
}
