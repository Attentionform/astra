/** Monospace face applied to code-block runs. */
import { toError } from '../errors';

export const COMPOSE_CODE_FONT = 'Courier New';
/** Header-row cell fill (6-digit hex, optional leading `#` accepted).
 * Legacy built-in for guideless composes; a guided compose emits the guide
 * `tables.headerFill` instead (see compileMarkdown). */
export const COMPOSE_TABLE_HEADER_FILL = 'D9D9D9';
/** Fence languages treated as chart requests (out of scope, fail closed). */
export const CHART_FENCE_LANGUAGES = new Set(['chart', 'mermaid']);

export const FOOTNOTE_DEF_PATTERN = /^\s{0,3}\[\^([^\]]+)\]:\s*(.*)$/;
export const FOOTNOTE_REF_PATTERN = /\[\^([^\]]+)\]/g;
export const CITATION_DEF_PATTERN = /^\s{0,3}\[@([^\]]+)\]:\s*(.*)$/;
export const CITATION_REF_PATTERN = /\[@([^\]]+)\]/g;
export const CITATION_KEY_PATTERN = /^[A-Za-z0-9][A-Za-z0-9:_-]*$/;
export const IMAGE_PATTERN = /!\[[^\]]*\]/;
export const HTML_PATTERN = /<[A-Za-z][A-Za-z0-9-]*(?:\s[^<>]*)?\/?>|<\/[A-Za-z][A-Za-z0-9-]*\s*>|<!--/;
export const AUTOLINK_PATTERN = /<[A-Za-z][A-Za-z0-9+.-]*:[^<>\s]+>|<[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}>/;
export const HEADING_ATTR_PATTERN = /\{[#.][^}]*\}\s*$/;
export const CONTAINER_FENCE_PATTERN = /^\s{0,3}:::/;
export const ENDNOTE_PREFIX = 'en-';
export const CHART_JSON_SUFFIX = '.json';
export const ALLOWED_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tif', 'tiff', 'svg']);
export const ALLOWED_LINK_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'ftp:', 'tel:', 'sms:']);
/** Dots per inch assumed when mapping image pixels to EMUs (Word web default). */
export const COMPOSE_IMAGE_DPI = 96;
/** EMUs per pixel at the assumed DPI (914400 EMUs per inch). */
export const COMPOSE_EMU_PER_PX = 914400 / COMPOSE_IMAGE_DPI;
/**
 * Default image width cap (6 inches). The blank template content box is
 * 6.5in wide (8.5in page minus 1in margins each side); 6in matches the
 * engine fallback width and always fits with margin to spare.
 */
export const COMPOSE_IMAGE_MAX_WIDTH_EMU = 5486400;
/** Default image height cap (9 inches, the blank content-box height). */
export const COMPOSE_IMAGE_MAX_HEIGHT_EMU = 8229600;
/** Inline data-URL images carry their own bytes, so no allowed directory applies. */
export const DATA_URL_IMAGE_PATTERN = /^data:image\/([\w.+-]+);base64,/i;
export const DATA_URL_BASE64_CHARS = /^[A-Za-z0-9+/=\s]+$/;
/** Decoded-byte cap for inline data-URL images (mirrors the file-asset cap). */
export const DATA_URL_MAX_BYTES = 15_000_000;
/**
 * Raw Markdown input ceiling in UTF-8 bytes, enforced before the lexer runs
 * so oversize manuscripts fail closed instead of expanding lexer state.
 * Three megabytes holds long manuscripts with headroom while keeping a
 * single compile bounded.
 */
export const COMPOSE_MARKDOWN_MAX_BYTES = 3_000_000;
/** Alert tags accepted on blockquotes, mapped to callout themes.
 * Legacy per-tag map for guideless composes; a guided compose takes the
 * guide `callouts.defaultTheme` instead (Markdown has no per-alert theme
 * syntax, so every tag is generic). */
export const ALERT_TYPE_TO_THEME: Record<string, string> = {
  NOTE: 'indigo',
  TIP: 'emerald',
  IMPORTANT: 'slate',
  WARNING: 'amber',
  CAUTION: 'rose',
};
export const ALERT_MARKER_PATTERN = /^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\](.*)$/i;
export const ALERT_TITLE_PATTERN = /^\*\*(.+)\*\*$/;
/** Exact seed texts of the blank template placeholder paragraphs. */
export const BLANK_SEED_TITLE = 'Blank Document';
export const BLANK_SEED_HINT = 'Start typing your content here...';
export const TEMPLATE_SEEDS: Record<string, string> = {
  blank: 'Blank Document',
  report: 'Q3 Executive',
  memo: 'MEMORANDUM',
};

export function fail(line: number, detail: string, remedy: string): never {
  throw toError('E_VALIDATION', `line ${line}: ${detail}`, remedy);
}

export function isEndnoteLabel(label: string): boolean {
  return label.startsWith(ENDNOTE_PREFIX);
}
