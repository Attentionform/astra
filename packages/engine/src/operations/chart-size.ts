/**
 * Chart drawing size: defaults, units, limits and layout warnings.
 *
 * Precedence for every dimension, highest first: the operation (`widthIn`/`widthEmu`, `heightIn`/`heightEmu`),
 * then the style guide's `Charts` section, then the built-in default. Each dimension is resolved on its own.
 *
 * Limits are ASTra's supported limits (0.1 in to 22 in): 22 in is the largest page Word can show and Word clips
 * a chart that is larger than the page, so a larger size is refused rather than written.
 */
import type { StyleGuideConfig } from '../style-guide/types';

export const EMU_PER_INCH = 914_400;
/** 0.1 in. */
export const CHART_MIN_EMU = 91_440;
/** 22 in. */
export const CHART_MAX_EMU = 20_116_800;

export const DEFAULT_CHART_WIDTH_EMU = 6 * EMU_PER_INCH; // 5,486,400
/** 2.6 in: a 2D chart plus a few lines of text fits on a Letter page with Word and LibreOffice. */
export const DEFAULT_CHART_HEIGHT_2D_EMU = Math.round(2.6 * EMU_PER_INCH); // 2,377,440
/** 3.5 in: unchanged from 0.5.2, to be reviewed on its own. */
export const DEFAULT_CHART_HEIGHT_3D_EMU = Math.round(3.5 * EMU_PER_INCH); // 3,200,400

export interface ChartSizeInput {
  widthEmu?: number;
  heightEmu?: number;
  widthIn?: number;
  heightIn?: number;
}

export interface ResolvedChartSize {
  widthEmu: number;
  heightEmu: number;
}

function sizeError(message: string, remedy: string): Error {
  return Object.assign(new Error(`E_VALIDATION: ${message}. Remedy: ${remedy}`), {
    code: 'E_VALIDATION',
    retryable: false,
    remedy,
  });
}

function toEmu(
  label: 'width' | 'height',
  emu: unknown,
  inches: unknown
): number | undefined {
  if (emu !== undefined && inches !== undefined) {
    throw sizeError(
      `chart ${label} is given twice (${label}Emu and ${label}In)`,
      `pass either ${label}Emu or ${label}In for the ${label}, not both`
    );
  }
  let value: number;
  let source: string;
  if (emu !== undefined) {
    source = `${label}Emu`;
    if (typeof emu !== 'number' || !Number.isFinite(emu)) {
      throw sizeError(`chart ${source} must be a finite number`, `pass ${source} as a number of EMU (914400 per inch) or use ${label}In`);
    }
    value = Math.round(emu);
  } else if (inches !== undefined) {
    source = `${label}In`;
    if (typeof inches !== 'number' || !Number.isFinite(inches)) {
      throw sizeError(`chart ${source} must be a finite number`, `pass ${source} as a number of inches, for example 6`);
    }
    value = Math.round(inches * EMU_PER_INCH);
  } else {
    return undefined;
  }
  if (value < CHART_MIN_EMU || value > CHART_MAX_EMU) {
    throw sizeError(
      `chart ${source} is outside the supported range (got ${emu !== undefined ? `${value} EMU` : `${inches} in`})`,
      `use ${label}In from 0.1 to 22 inches (${label}Emu from ${CHART_MIN_EMU} to ${CHART_MAX_EMU}); a chart larger than a page is clipped by Word`
    );
  }
  return value;
}

/** Validates the dimensions an operation gives; returns only the dimensions that were given, in EMU. */
export function parseChartSize(input: ChartSizeInput): { widthEmu?: number; heightEmu?: number } {
  const widthEmu = toEmu('width', input.widthEmu, input.widthIn);
  const heightEmu = toEmu('height', input.heightEmu, input.heightIn);
  return {
    ...(widthEmu !== undefined ? { widthEmu } : {}),
    ...(heightEmu !== undefined ? { heightEmu } : {}),
  };
}

/** Guide `Charts` size as EMU, or undefined when the guide gives no usable value. */
function guideEmu(inches: unknown): number | undefined {
  if (typeof inches !== 'number' || !Number.isFinite(inches)) return undefined;
  const emu = Math.round(inches * EMU_PER_INCH);
  return emu >= CHART_MIN_EMU && emu <= CHART_MAX_EMU ? emu : undefined;
}

/**
 * The size of a new chart: each dimension is the operation's value, else the guide's, else the built-in default.
 * Giving only the width does not scale the height.
 */
export function resolveNewChartSize(input: ChartSizeInput, is3D: boolean, guide?: StyleGuideConfig, area?: UsableArea): ResolvedChartSize {
  const given = parseChartSize(input);
  const charts = guide?.charts;
  const guideWidth = guideEmu(is3D ? charts?.width3dIn ?? charts?.widthIn : charts?.widthIn);
  const guideHeight = guideEmu(is3D ? charts?.height3dIn : charts?.heightIn);
  // The built-in default follows the page: it is never larger than the text area it is placed in. A size that the operation or the
  // guide gives is kept as it is (a layout warning tells the caller when it does not fit).
  const fit = (value: number, room: number | undefined): number => (room !== undefined && room > 0 ? Math.max(CHART_MIN_EMU, Math.min(value, room)) : value);
  return {
    widthEmu: given.widthEmu ?? guideWidth ?? fit(DEFAULT_CHART_WIDTH_EMU, area?.widthEmu),
    heightEmu: given.heightEmu ?? guideHeight ?? fit(is3D ? DEFAULT_CHART_HEIGHT_3D_EMU : DEFAULT_CHART_HEIGHT_2D_EMU, area?.heightEmu),
  };
}

export interface UsableArea {
  widthEmu: number;
  heightEmu: number;
}

/**
 * An inline chart wider than the text width, or taller than the page's usable height, is clipped by Word.
 * The warning is advisory: the chart is still written. Floating charts are not checked.
 */
export function chartLayoutWarning(
  size: ResolvedChartSize,
  area: UsableArea | undefined,
  floating: boolean,
  widthName: 'text width' | 'column width' = 'text width'
): string | undefined {
  if (floating || !area) return undefined;
  const parts: string[] = [];
  if (size.widthEmu > area.widthEmu) {
    parts.push(`${fmtIn(size.widthEmu)} in wide, the ${widthName} is ${fmtIn(area.widthEmu)} in`);
  }
  if (size.heightEmu > area.heightEmu) {
    parts.push(`${fmtIn(size.heightEmu)} in tall, the usable page height is ${fmtIn(area.heightEmu)} in`);
  }
  if (parts.length === 0) return undefined;
  return `layoutWarning: the chart is ${parts.join(' and ')}; Word clips an inline chart at the page edge. Use a smaller widthIn or heightIn.`;
}

function fmtIn(emu: number): string {
  return String(Math.round((emu / EMU_PER_INCH) * 100) / 100);
}

/** Page sizes in twips (portrait) for the paper names a section can carry. */
const PAPER_TWIPS: Record<string, { width: number; height: number }> = {
  Letter: { width: 12240, height: 15840 },
  A4: { width: 11906, height: 16838 },
  Legal: { width: 12240, height: 20160 },
  A3: { width: 16838, height: 23811 },
};

/**
 * Text width and usable page height in EMU for a section's page setup. The size comes from the section's own
 * `w:pgSz` when it is known (any paper, any custom size, already in the orientation in use); otherwise from the
 * paper name (Letter when there is none), turned to the orientation. Margins default to 1 in.
 */
export function usableAreaOf(pageSetup: {
  orientation?: string;
  paperSize?: string;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  gutter?: number;
  /** True when the gutter is on the top edge (a setting of the document); it is on the left edge otherwise. */
  gutterAtTop?: boolean;
  pageWidthTwips?: number;
  pageHeightTwips?: number;
} | null | undefined): UsableArea {
  let widthTw: number;
  let heightTw: number;
  if (pageSetup?.pageWidthTwips && pageSetup.pageWidthTwips > 0 && pageSetup.pageHeightTwips && pageSetup.pageHeightTwips > 0) {
    widthTw = pageSetup.pageWidthTwips;
    heightTw = pageSetup.pageHeightTwips;
  } else {
    const paper = PAPER_TWIPS[pageSetup?.paperSize ?? 'Letter'] ?? PAPER_TWIPS.Letter;
    widthTw = paper.width;
    heightTw = paper.height;
    if (pageSetup?.orientation === 'landscape') [widthTw, heightTw] = [heightTw, widthTw];
  }
  const left = pageSetup?.marginLeft ?? 1440;
  const right = pageSetup?.marginRight ?? 1440;
  const top = pageSetup?.marginTop ?? 1440;
  const bottom = pageSetup?.marginBottom ?? 1440;
  // The gutter is binding space: it comes off the width when it is on the left, and off the height when it is on the top.
  const gutter = pageSetup?.gutter ?? 0;
  const gutterWidth = pageSetup?.gutterAtTop ? 0 : gutter;
  const gutterHeight = pageSetup?.gutterAtTop ? gutter : 0;
  const twipsToEmu = EMU_PER_INCH / 1440;
  return {
    widthEmu: Math.max(0, Math.round((widthTw - left - right - gutterWidth) * twipsToEmu)),
    heightEmu: Math.max(0, Math.round((heightTw - top - bottom - gutterHeight) * twipsToEmu)),
  };
}

/** The Letter content box that the image cap was written for: 6.5 in wide and 9 in high. */
export const LETTER_CONTENT_BOX: UsableArea = { widthEmu: 5943600, heightEmu: 8229600 };

/** The default space between columns, in twips: Word's, used when the file does not give one. */
const DEFAULT_COLUMN_SPACE_TWIPS = 720;

type ColumnPage = Parameters<typeof usableAreaOf>[0] & { columns?: number; columnSpace?: number; columnWidths?: number[]; columnsUnknown?: boolean };

/**
 * The width of the narrowest text column of a section in EMU, and whether the columns are of different widths or not known.
 * Equal columns: the text width less the real space between them (`w:space`, 0.5 in when the file gives none), shared by the count.
 * Listed columns of different widths: the narrowest, so that content placed in any of them fits.
 * A listing without every width: the equal share as a best guess, flagged as not known.
 */
export function columnLayoutOf(pageSetup: ColumnPage | null | undefined): { widthEmu: number; unequal: boolean; unknown: boolean } | undefined {
  if (!pageSetup || !pageSetup.pageWidthTwips) return undefined;
  const columns = pageSetup.columns && pageSetup.columns > 1 ? pageSetup.columns : 1;
  if (columns === 1) return undefined;
  const twipsToEmu = EMU_PER_INCH / 1440;
  const area = usableAreaOf(pageSetup);
  const widths = pageSetup.columnWidths;
  if (widths && widths.length === columns) {
    const narrowest = Math.min(...widths);
    const widest = Math.max(...widths);
    return { widthEmu: Math.max(0, Math.floor(narrowest * twipsToEmu)), unequal: widest - narrowest > Math.max(1, widest * 0.01), unknown: false };
  }
  const gap = (pageSetup.columnSpace ?? DEFAULT_COLUMN_SPACE_TWIPS) * twipsToEmu * (columns - 1);
  return { widthEmu: Math.max(0, Math.floor((area.widthEmu - gap) / columns)), unequal: false, unknown: pageSetup.columnsUnknown === true };
}

/**
 * A note for the caller when content placed in a section with columns of different or unknown widths was sized to the narrowest
 * column, because the engine cannot tell which column the content lands in. Undefined for one column and for equal columns.
 */
export function columnLayoutNote(pageSetup: ColumnPage | null | undefined, what: string): string | undefined {
  const layout = columnLayoutOf(pageSetup);
  if (!layout || (!layout.unequal && !layout.unknown)) return undefined;
  const inches = Math.round((layout.widthEmu / EMU_PER_INCH) * 100) / 100;
  return `${what} was sized to ${inches} in, the ${layout.unknown ? 'estimated' : 'narrowest'} column of a section whose columns ${layout.unknown ? 'are not all given a width' : 'have different widths'}; give an explicit size if it goes into a wider column`;
}

/**
 * The box that content placed in a section has to fit: the usable area, narrowed to one text column when the section has more than one
 * (see `columnLayoutOf`). A section without a page size gets the Letter box.
 */
export function contentBoxOf(pageSetup: ColumnPage | null | undefined): UsableArea {
  if (!pageSetup || !pageSetup.pageWidthTwips) return LETTER_CONTENT_BOX;
  const area = usableAreaOf(pageSetup);
  const layout = columnLayoutOf(pageSetup);
  return layout ? { widthEmu: layout.widthEmu, heightEmu: area.heightEmu } : area;
}

/** Letter's text width in twips: 6.5 in. What an engine-written table grid or tab stop used before the page could differ. */
export const LETTER_TEXT_WIDTH_TWIPS = 9360;

/** The text width of a section in twips (one column's width when the section has columns). A section without a page size gets Letter's. */
export function textWidthTwipsOf(pageSetup: Parameters<typeof contentBoxOf>[0]): number {
  if (!pageSetup || !pageSetup.pageWidthTwips) return LETTER_TEXT_WIDTH_TWIPS;
  const twips = Math.round(contentBoxOf(pageSetup).widthEmu / (EMU_PER_INCH / 1440));
  return Math.max(1, twips);
}
