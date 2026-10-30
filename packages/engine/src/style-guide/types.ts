/**
 * Style guide schema — typed design system for machine-built documents.
 *
 * A guide is authored as Markdown (`style-guide.md`), parsed by
 * `./parser` into a `StyleGuideConfig`, and (from a later slice) applied
 * by `OperationEngine` as omitted-param defaults plus advisory linting.
 */
export interface StyleGuidePalette {
  primary: string; // 6-char uppercase hex, no leading `#` (e.g. `1E3A8A`)
  secondary?: string;
  accent?: string;
  neutralDark?: string;
  neutralLight?: string;
  chartPalette: string[]; // at least 4 hex colors, series order
}

export interface StyleGuideTables {
  headerFill?: string;
  headerText?: string; // contrast text color for header cells
  headerBold?: boolean; // bold header row text (default true)
  alternatingFill?: string; // zebra-striping fill
  borderColor?: string;
  keepTogether?: boolean; // cantSplit rows + tblHeader/keepNext header
}

export interface StyleGuideTypography {
  fontFamily?: string;
  codeFont?: string;
  heading1?: StyleGuideHeading;
  heading2?: StyleGuideHeading;
}

/** Heading style from `- Heading 1: #1E3A8A, bold` (color + bold only). */
export interface StyleGuideHeading {
  color?: string; // 6-char uppercase hex, no leading `#` (e.g. `1E3A8A`)
  bold?: boolean;
}

export const STYLE_GUIDE_CALLOUT_THEMES = [
  'indigo',
  'emerald',
  'amber',
  'slate',
  'rose',
] as const;

export type StyleGuideCalloutTheme =
  (typeof STYLE_GUIDE_CALLOUT_THEMES)[number];

export type StyleGuideCalloutTag =
  'note' | 'tip' | 'important' | 'warning' | 'caution';

export interface StyleGuideCallouts {
  defaultTheme?: StyleGuideCalloutTheme;
  preferredThemes?: StyleGuideCalloutTheme[];
  tagThemes?: Partial<Record<StyleGuideCalloutTag, StyleGuideCalloutTheme>>;
}

export interface StyleGuideCodeBlocks {
  background?: string;
  borderColor?: string;
}

export const STYLE_GUIDE_MARGINS = ['standard', 'narrow', 'wide'] as const;

export type StyleGuideMargin = (typeof STYLE_GUIDE_MARGINS)[number];

export interface StyleGuideLayout {
  margin?: StyleGuideMargin;
  preventOrphanHeadings?: boolean;
  widowControl?: boolean;
}

export interface StyleGuideConfig {
  name: string;
  version: 1; // guide format version — parser warns (never throws) on unknown
  description?: string;
  palette: StyleGuidePalette;
  tables: StyleGuideTables;
  typography: StyleGuideTypography;
  callouts: StyleGuideCallouts;
  codeBlocks: StyleGuideCodeBlocks;
  layout: StyleGuideLayout;
}

/** One human-readable diagnostic; value problems name their 1-based line. */
export type StyleGuideWarning = string;

export interface StyleGuideParseResult {
  config: StyleGuideConfig;
  warnings: StyleGuideWarning[];
}
