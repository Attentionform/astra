/**
 * The colours of a style guide: the main colour, optional secondary and accent colours, dark and light neutrals, and the colours for chart series, in order.
 */
export interface StyleGuidePalette {
  primary: string; // 6-char uppercase hex, no leading `#` (e.g. `1E3A8A`)
  secondary?: string;
  accent?: string;
  neutralDark?: string;
  neutralLight?: string;
  chartPalette: string[]; // at least 4 hex colors, series order
}

/**
 * The table settings of a style guide: header fill and text, alternating row fill, borders, whether rows stay together, and the space after a table.
 */
export interface StyleGuideTables {
  headerFill?: string;
  headerText?: string; // contrast text color for header cells
  headerBold?: boolean; // bold header row text (default true)
  alternatingFill?: string; // zebra-striping fill
  borderColor?: string;
  keepTogether?: boolean; // cantSplit rows + tblHeader/keepNext header
  /** Space before the first block a compose call creates after a table, in points (0 to 72; built-in 12). */
  spaceAfterPt?: number;
}

/**
 * The chart defaults of a style guide: sizes in inches, text sizes in points, the gridline colour and the font. A key that is missing means that the built-in default applies.
 */
export interface StyleGuideCharts {
  widthIn?: number;
  heightIn?: number;
  width3dIn?: number;
  height3dIn?: number;
  gridlineColor?: string; // 6-char uppercase hex
  fontFamily?: string;
  fontSizePt?: number;
  titleSizePt?: number;
}

/**
 * The type settings of a style guide: the font, the code font, and the look of the first two heading levels. They are checked and not drawn.
 */
export interface StyleGuideTypography {
  fontFamily?: string;
  codeFont?: string;
  heading1?: StyleGuideHeading;
  heading2?: StyleGuideHeading;
}

/**
 * The look of a heading in a style guide: its colour, and whether it is bold.
 */
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

/**
 * A callout colour theme.
 */
export type StyleGuideCalloutTheme =
  (typeof STYLE_GUIDE_CALLOUT_THEMES)[number];

/**
 * An alert tag that Markdown can use, as in `> [!NOTE]`.
 */
export type StyleGuideCalloutTag =
  'note' | 'tip' | 'important' | 'warning' | 'caution';

/**
 * The callout settings of a style guide: the default theme, the preferred themes, a theme for each alert tag, and the space after a callout.
 */
export interface StyleGuideCallouts {
  defaultTheme?: StyleGuideCalloutTheme;
  preferredThemes?: StyleGuideCalloutTheme[];
  tagThemes?: Partial<Record<StyleGuideCalloutTag, StyleGuideCalloutTheme>>;
  /** Space before the first block a compose call creates after a callout, in points (0 to 72; built-in 8). */
  spaceAfterPt?: number;
}

/**
 * The code block settings of a style guide: the background and the border colour.
 */
export interface StyleGuideCodeBlocks {
  background?: string;
  borderColor?: string;
}

export const STYLE_GUIDE_MARGINS = ['standard', 'narrow', 'wide'] as const;

/**
 * A margin size that a style guide can ask for.
 */
export type StyleGuideMargin = (typeof STYLE_GUIDE_MARGINS)[number];

/**
 * The layout settings of a style guide: the margin, and the control of lonely headings and widow lines. They are checked and not drawn.
 */
export interface StyleGuideLayout {
  margin?: StyleGuideMargin;
  preventOrphanHeadings?: boolean;
  widowControl?: boolean;
}

/**
 * A style guide after it has been read: its name, a palette and settings for tables, charts, typography, callouts, code blocks and layout.
 */
export interface StyleGuideConfig {
  name: string;
  version: 1; // guide format version — parser warns (never throws) on unknown
  description?: string;
  palette: StyleGuidePalette;
  tables: StyleGuideTables;
  /** Optional: absent when the guide has no Charts section. */
  charts?: StyleGuideCharts;
  typography: StyleGuideTypography;
  callouts: StyleGuideCallouts;
  codeBlocks: StyleGuideCodeBlocks;
  layout: StyleGuideLayout;
}

/**
 * One note from reading a style guide. A note about a value names its line, counted from 1.
 */
export type StyleGuideWarning = string;

/**
 * What `parseStyleGuide` returns: the `config`, and `warnings` for what it could not use.
 */
export interface StyleGuideParseResult {
  config: StyleGuideConfig;
  warnings: StyleGuideWarning[];
}
