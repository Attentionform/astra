import { toError } from '../errors';
import { LENGTH_UNITS, parseLength, splitLength, twipsToInches, type LengthInput } from './length';
import { PAPER_NAMES, detectPaper, findPaper, type Paper } from './papers';
import { LOW_MARGIN_WARNING, MAX_PAGE_SIDE, MIN_PAGE_SIDE, MIN_TEXT_HEIGHT, MIN_TEXT_WIDTH } from './limits';

export { MIN_PAGE_SIDE, MAX_PAGE_SIDE, MIN_TEXT_WIDTH, MIN_TEXT_HEIGHT, LOW_MARGIN_WARNING };

/**
 * A page size as a caller writes it: a paper name (`"A4"`, case does not matter), a width and height (`{ width: "148mm", height: "210mm" }`),
 * or a shorthand string (`"148x210mm"`, `"6x9in"`, `"210mm x 297mm"`).
 */
export type PageSizeInput = string | { width: LengthInput; height: LengthInput };

/** Margins as a caller writes them: one length for all four page margins, or an object with any of the fields. */
export type MarginsInput =
  | LengthInput
  | {
      top?: LengthInput;
      bottom?: LengthInput;
      left?: LengthInput;
      right?: LengthInput;
      /** The distance from the page edge to the header. */
      header?: LengthInput;
      /** The distance from the page edge to the footer. */
      footer?: LengthInput;
      /** Extra space on the binding edge. */
      gutter?: LengthInput;
    };

export interface ResolvedPageSize {
  /** Width and height in whole twips, as given (portrait or landscape is the caller's). */
  width: number;
  height: number;
  /** The named paper that these dimensions match, or undefined for a custom size. */
  paper?: Paper;
  /** True when the caller gave a paper name, whose dimensions are portrait; false when the caller gave the dimensions. */
  named: boolean;
}

export type MarginName = 'top' | 'bottom' | 'left' | 'right' | 'header' | 'footer' | 'gutter';
const MARGIN_NAMES: readonly MarginName[] = ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'];

function shorthand(text: string, what: string): { width: number; height: number } | null {
  const halves = text.split(/\s*[x×]\s*/i);
  if (halves.length !== 2) return null;
  const a = splitLength(halves[0]);
  const b = splitLength(halves[1]);
  if (!a || !b) return null;
  // "6x9in": the unit on the second number applies to the first.
  const unitA = a.unit === '' && b.unit !== '' ? b.unit : a.unit;
  return { width: parseLength(`${a.value}${unitA}`, `${what} width`), height: parseLength(`${b.value}${b.unit}`, `${what} height`) };
}

/**
 * Resolves a page size to whole twips. A name is looked up in the paper table; anything else must be a width and a height.
 * Does not check limits; `checkPage` does that once the margins are known.
 */
export function resolvePageSize(input: unknown, what = 'paperSize'): ResolvedPageSize {
  let width: number;
  let height: number;
  if (typeof input === 'string') {
    const paper = findPaper(input);
    if (paper) {
      return { width: paper.width, height: paper.height, paper, named: true };
    }
    const parsed = shorthand(input, what);
    if (!parsed) {
      throw toError(
        'E_VALIDATION',
        `${what} '${input}' is not a known paper or a size`,
        `use one of ${PAPER_NAMES.join(', ')}, or give a size such as "148x210mm" or { "width": "6in", "height": "9in" }; units are ${LENGTH_UNITS.join(', ')}`
      );
    }
    ({ width, height } = parsed);
  } else if (input !== null && typeof input === 'object' && !Array.isArray(input)) {
    const o = input as Record<string, unknown>;
    const extra = Object.keys(o).filter((k) => k !== 'width' && k !== 'height');
    if (extra.length > 0 || o.width === undefined || o.height === undefined) {
      throw toError('E_VALIDATION', `${what} needs exactly a width and a height${extra.length > 0 ? `; unknown field ${extra.join(', ')}` : ''}`, `pass { "width": "148mm", "height": "210mm" } or a paper name`);
    }
    width = parseLength(o.width, `${what} width`);
    height = parseLength(o.height, `${what} height`);
  } else {
    throw toError('E_VALIDATION', `${what} must be a paper name or a size`, `use one of ${PAPER_NAMES.join(', ')}, or give { "width": "148mm", "height": "210mm" }`);
  }
  return { width, height, paper: detectPaper(width, height), named: false };
}

/**
 * Resolves margins to whole twips. One length means the four page margins. An object holds any of the seven fields. A field
 * that is not given is absent from the result, so that a caller can leave it as the document has it.
 */
export function resolveMargins(input: unknown, what = 'margins'): Partial<Record<MarginName, number>> {
  if (input === undefined) return {};
  if (typeof input === 'number' || typeof input === 'string') {
    const all = parseLength(input, what);
    return { top: all, bottom: all, left: all, right: all };
  }
  if (input === null || typeof input !== 'object' || Array.isArray(input)) {
    throw toError('E_VALIDATION', `${what} must be a length or an object of lengths`, `pass "25mm" for every margin, or { "top": "20mm", "left": "1in" }`);
  }
  const o = input as Record<string, unknown>;
  const unknown = Object.keys(o).filter((k) => !(MARGIN_NAMES as readonly string[]).includes(k));
  if (unknown.length > 0) {
    throw toError('E_VALIDATION', `${what} has an unknown field: ${unknown.join(', ')}`, `use ${MARGIN_NAMES.join(', ')}`);
  }
  const out: Partial<Record<MarginName, number>> = {};
  for (const key of MARGIN_NAMES) if (o[key] !== undefined) out[key] = parseLength(o[key], `${what}.${key}`);
  return out;
}

export interface PageGeometry {
  width: number;
  height: number;
  top: number;
  bottom: number;
  left: number;
  right: number;
  header?: number;
  footer?: number;
  gutter?: number;
  /** True when the gutter is at the top edge (`w:gutterAtTop`); it is on the left edge otherwise. */
  gutterAtTop?: boolean;
}

/** What a request touches, so that a check can leave the rest of an existing page alone. */
export interface PageCheckScope {
  /** True when the request sets the page size (a paper name or a width and a height). A turn of orientation is not a size. */
  size?: boolean;
  /** The margin fields that the request sets. Untouched fields keep their value and are not an error, even a negative top or bottom margin that Word accepts. */
  margins?: readonly MarginName[];
  /** The page as it is now. A text area that is already under the minimum is not refused unless the request makes it smaller. */
  before?: PageGeometry;
}

const textAreaOf = (g: PageGeometry): { width: number; height: number } => {
  const gutter = g.gutter ?? 0;
  return { width: g.width - g.left - g.right - (g.gutterAtTop ? 0 : gutter), height: g.height - g.top - g.bottom - (g.gutterAtTop ? gutter : 0) };
};

/**
 * Checks that a page is one that can be laid out. Throws `E_VALIDATION` for a size or margin that is out of range or leaves too little
 * text area. Returns warnings for what is valid but likely to be a mistake. `where` names the section in a message.
 *
 * Without `scope` every value is checked, as for a new document. With it, the checks on single values (size range, sign and whole numbers)
 * apply only to what the request sets, and the text area is checked on the resulting page: a page that was already under the minimum is
 * refused only when the request makes the area smaller.
 */
export function checkPage(g: PageGeometry, where = 'page setup', scope?: PageCheckScope): string[] {
  const fail = (message: string, remedy: string): never => {
    throw toError('E_VALIDATION', `${where}: ${message}`, remedy);
  };
  const touchedMargins: readonly MarginName[] = scope?.margins ?? MARGIN_NAMES;
  if (scope?.size !== false) {
    for (const [name, value] of [['width', g.width], ['height', g.height]] as const) {
      if (!Number.isInteger(value) || value < MIN_PAGE_SIDE || value > MAX_PAGE_SIDE) {
        fail(`page ${name} is ${Number.isFinite(value) ? `${twipsToInches(value)} in` : String(value)}, which is outside 1 in to 22 in`, `give a ${name} from 1 in to 22 in, for example "${name === 'width' ? '148mm' : '210mm'}". Word cannot read a larger page and the file validator rejects it`);
      }
    }
  }
  for (const name of touchedMargins) {
    const value = g[name];
    if (value === undefined) continue;
    if (!Number.isInteger(value) || value < 0) {
      fail(`${name} must be zero or more, got ${String(value)} twips`, `give ${name} as a length of zero or more, for example "25mm"`);
    }
  }
  const gutter = g.gutter ?? 0;
  const after = textAreaOf(g);
  const before = scope?.before ? textAreaOf(scope.before) : undefined;
  if (after.width < MIN_TEXT_WIDTH && (!before || after.width < before.width)) {
    fail(`the left and right margins${gutter > 0 && !g.gutterAtTop ? ' and gutter' : ''} leave ${twipsToInches(Math.max(after.width, 0))} in of text width on a ${twipsToInches(g.width)} in wide page; at least 0.75 in is needed`, 'make the margins smaller or the page wider');
  }
  if (after.height < MIN_TEXT_HEIGHT && (!before || after.height < before.height)) {
    fail(`the top and bottom margins${gutter > 0 && g.gutterAtTop ? ' and gutter' : ''} leave ${twipsToInches(Math.max(after.height, 0))} in of text height on a ${twipsToInches(g.height)} in high page; at least 0.5 in is needed`, 'make the margins smaller or the page taller');
  }
  const warnings: string[] = [];
  const low = (['top', 'bottom', 'left', 'right'] as const).filter((n) => touchedMargins.includes(n) && g[n] < LOW_MARGIN_WARNING && g[n] >= 0);
  if (low.length > 0) warnings.push(`${where}: the ${low.join(', ')} margin is under 0.25 in, which a printer may clip`);
  if (g.header !== undefined && g.header > g.top && (touchedMargins.includes('header') || touchedMargins.includes('top'))) warnings.push(`${where}: the header distance is larger than the top margin, so the header pushes the body down`);
  if (g.footer !== undefined && g.footer > g.bottom && (touchedMargins.includes('footer') || touchedMargins.includes('bottom'))) warnings.push(`${where}: the footer distance is larger than the bottom margin, so the footer pushes the body up`);
  return warnings;
}
