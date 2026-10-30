/**
 * Built-in fallback values for style guide parsing.
 *
 * PRECEDENCE INVARIANT (global — governs engine defaulting, agent context,
 * and advisory linting; wired in a later slice inside `OperationEngine`,
 * stated here so every default below is read in the right order):
 *
 *   explicit op param > style guide value > built-in default (this file)
 *
 * Consequences for the later wiring: a caller-passed value always wins
 * (even when it clashes with the guide — clashes surface as advisory
 * warnings, never silent overrides); guide values fill only omitted params;
 * chart palettes fill only uncolored series in order without shifting the
 * indices of explicitly colored series.
 */
import type { StyleGuideConfig } from './types';

/** Corporate-neutral fallback; every field is valid OOXML-ready output. */
export const DEFAULT_STYLE_GUIDE: StyleGuideConfig = {
  name: 'Default Corporate',
  version: 1,
  description: 'Built-in corporate-neutral fallback applied when a guide omits a value.',
  palette: {
    primary: '1E3A8A',
    secondary: '0284C7',
    accent: 'D97706',
    neutralDark: '1E293B',
    neutralLight: 'F8FAFC',
    chartPalette: ['1E3A8A', '0284C7', '10B981', 'F59E0B', '6366F1', 'EC4899'],
  },
  tables: {
    headerFill: '1E3A8A',
    headerText: 'FFFFFF',
    headerBold: true,
    alternatingFill: 'F8FAFC',
    borderColor: 'CBD5E1',
    keepTogether: true,
  },
  typography: {
    fontFamily: 'Calibri',
    codeFont: 'Consolas',
  },
  callouts: {
    defaultTheme: 'slate',
    preferredThemes: ['slate', 'emerald', 'amber'],
  },
  codeBlocks: {
    background: 'F1F5F9',
    borderColor: 'CBD5E1',
  },
  layout: {
    margin: 'standard',
    preventOrphanHeadings: true,
    widowControl: true,
  },
};

/** Fresh deep copy — the parser mutates its working config freely. */
export function createDefaultStyleGuide(): StyleGuideConfig {
  return JSON.parse(JSON.stringify(DEFAULT_STYLE_GUIDE)) as StyleGuideConfig;
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const nested of Object.values(value)) deepFreeze(nested);
  }
  return value;
}

deepFreeze(DEFAULT_STYLE_GUIDE);
