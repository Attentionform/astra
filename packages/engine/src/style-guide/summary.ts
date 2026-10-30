/**
 * Compact session-facing summary of an active style guide.
 *
 * The agent context layer (MCP `inspect_document`, CLI `validate`) carries
 * this shape instead of the full config: enough to name the house style
 * (header colors, chart order, callout theme) without bloating context.
 * Pure projection — never mutates the guide, never throws.
 */
import type { StyleGuideCalloutTag, StyleGuideCalloutTheme, StyleGuideConfig } from './types';

export interface StyleGuideSummary {
  name: string;
  headerFill?: string;
  headerText?: string;
  chartPalette: string[];
  calloutTheme?: string;
  calloutTagThemes?: Partial<Record<StyleGuideCalloutTag, StyleGuideCalloutTheme>>;
}

/** Projects a resolved guide to its compact agent-context summary. */
export function summarizeStyleGuide(guide: StyleGuideConfig): StyleGuideSummary {
  const palette = Array.isArray(guide?.palette?.chartPalette)
    ? guide.palette.chartPalette.filter((c): c is string => typeof c === 'string' && c !== '')
    : [];
  const out: StyleGuideSummary = {
    name: typeof guide?.name === 'string' && guide.name !== '' ? guide.name : 'Style Guide',
    chartPalette: [...palette],
  };
  const headerFill = guide?.tables?.headerFill;
  if (typeof headerFill === 'string' && headerFill !== '') out.headerFill = headerFill;
  const headerText = guide?.tables?.headerText;
  if (typeof headerText === 'string' && headerText !== '') out.headerText = headerText;
  const calloutTheme = guide?.callouts?.defaultTheme;
  if (calloutTheme !== undefined && calloutTheme !== null && String(calloutTheme) !== '') {
    out.calloutTheme = String(calloutTheme);
  }
  const tagThemes = guide?.callouts?.tagThemes;
  if (tagThemes !== undefined && tagThemes !== null && typeof tagThemes === 'object') {
    out.calloutTagThemes = { ...tagThemes };
  }
  return out;
}
