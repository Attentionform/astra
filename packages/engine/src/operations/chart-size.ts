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
export function resolveNewChartSize(input: ChartSizeInput, is3D: boolean, guide?: StyleGuideConfig): ResolvedChartSize {
  const given = parseChartSize(input);
  const charts = guide?.charts;
  const guideWidth = guideEmu(is3D ? charts?.width3dIn ?? charts?.widthIn : charts?.widthIn);
  const guideHeight = guideEmu(is3D ? charts?.height3dIn : charts?.heightIn);
  return {
    widthEmu: given.widthEmu ?? guideWidth ?? DEFAULT_CHART_WIDTH_EMU,
    heightEmu: given.heightEmu ?? guideHeight ?? (is3D ? DEFAULT_CHART_HEIGHT_3D_EMU : DEFAULT_CHART_HEIGHT_2D_EMU),
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
  floating: boolean
): string | undefined {
  if (floating || !area) return undefined;
  const parts: string[] = [];
  if (size.widthEmu > area.widthEmu) {
    parts.push(`${fmtIn(size.widthEmu)} in wide, the text width is ${fmtIn(area.widthEmu)} in`);
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

/** Text width and usable page height in EMU for a section's page setup (Letter, 1 in margins when unknown). */
export function usableAreaOf(pageSetup: {
  orientation?: string;
  paperSize?: string;
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
} | null | undefined): UsableArea {
  const a4 = pageSetup?.paperSize === 'A4';
  let widthTw = a4 ? 11906 : 12240;
  let heightTw = a4 ? 16838 : 15840;
  if (pageSetup?.orientation === 'landscape') [widthTw, heightTw] = [heightTw, widthTw];
  const left = pageSetup?.marginLeft ?? 1440;
  const right = pageSetup?.marginRight ?? 1440;
  const top = pageSetup?.marginTop ?? 1440;
  const bottom = pageSetup?.marginBottom ?? 1440;
  const twipsToEmu = EMU_PER_INCH / 1440;
  return {
    widthEmu: Math.max(0, Math.round((widthTw - left - right) * twipsToEmu)),
    heightEmu: Math.max(0, Math.round((heightTw - top - bottom) * twipsToEmu)),
  };
}
