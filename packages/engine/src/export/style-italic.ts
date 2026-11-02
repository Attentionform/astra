/**
 * Italic of a run for the HTML and Markdown exports, from the run, its character style, its paragraph
 * style, its enclosing table style and docDefaults. A Subtitle or Quote paragraph, a styled table, or
 * a document default sets italic without `w:i` on individual runs.
 *
 * Italic is a toggle property (ECMA-376 17.7.3): direct formatting on the run (`w:i` on or off) decides
 * alone. Without it, the four style levels (docDefaults, table style, paragraph style and character style)
 * toggle each other, combining by exclusive or. Within one style, the closest definition in the
 * `w:basedOn` chain wins.
 *
 * The styles come from the parser that an export is given (`DocxParser.getStyleElements`), so a style
 * edited in a session and not yet saved is read as it is now. An export without a parser reads direct
 * formatting only.
 */
import type { DocumentAST } from '../ast';

const BASED_ON_LIMIT = 24;

/** The `w:style` elements of one export, and what has been worked out from them. */
interface StyleLookup {
  elements: Map<string, Element>;
  italic: Map<string, boolean>;
  docDefaults?: Element;
  docDefaultsItalic?: boolean;
}

const lookups = new WeakMap<DocumentAST, StyleLookup>();

/** Set (or, without a parser, clear) the styles that `effectiveRunItalic` reads for this AST. Call at the start of an export. */
export function useStylesForExport(ast: DocumentAST, parser: unknown): void {
  const source = parser as {
    getStyleElements?: () => Map<string, Element>;
    getDocDefaultsElement?: () => Element | undefined;
  } | null | undefined;
  if (source && typeof source.getStyleElements === 'function') {
    try {
      const elements = source.getStyleElements();
      const docDefaults = typeof source.getDocDefaultsElement === 'function'
        ? source.getDocDefaultsElement()
        : (elements.get('docDefaults')?.tagName === 'w:docDefaults' ? elements.get('docDefaults') : undefined);
      lookups.set(ast, { elements, italic: new Map(), docDefaults });
      return;
    } catch {
      /* fall through: an export never fails for want of styles */
    }
  }
  lookups.delete(ast);
}

function isOff(val: string | null | undefined): boolean {
  const v = (val ?? '').toLowerCase();
  return v === '0' || v === 'false' || v === 'off' || v === 'none';
}

function directChild(parent: Element | null | undefined, tag: string): Element | undefined {
  if (!parent) return undefined;
  const kids = parent.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const child = kids[i] as Element;
    if (child.nodeType === 1 && child.tagName === tag) return child;
  }
  return undefined;
}

/** `true` / `false` when the run properties say so, `undefined` when they do not mention italic. */
export function directItalic(rPr: Element | undefined | null): boolean | undefined {
  const i = directChild(rPr, 'w:i');
  if (!i) return undefined;
  return !isOff(i.getAttribute('w:val'));
}

/** Italic of one style: the closest `w:i` in its `w:basedOn` chain (`false` when none is set). */
export function italicOfStyle(lookup: StyleLookup | undefined, styleId: string | null | undefined): boolean {
  if (!lookup || !styleId) return false;
  const cached = lookup.italic.get(styleId);
  if (cached !== undefined) return cached;
  const seen = new Set<string>();
  let id: string | null | undefined = styleId;
  let result = false;
  for (let hop = 0; id && hop < BASED_ON_LIMIT && !seen.has(id); hop++) {
    seen.add(id);
    const def: Element | undefined = lookup.elements.get(id);
    if (!def) break;
    const own = directItalic(directChild(def, 'w:rPr'));
    if (own !== undefined) {
      result = own;
      break;
    }
    id = directChild(def, 'w:basedOn')?.getAttribute('w:val');
  }
  lookup.italic.set(styleId, result);
  return result;
}

/** Italic setting from `w:docDefaults/w:rPrDefault/w:rPr`, if any. */
export function italicOfDocDefaults(lookup: StyleLookup | undefined): boolean {
  if (!lookup) return false;
  if (lookup.docDefaultsItalic !== undefined) return lookup.docDefaultsItalic;
  const defaults = lookup.docDefaults ?? (lookup.elements.get('docDefaults')?.tagName === 'w:docDefaults' ? lookup.elements.get('docDefaults') : undefined);
  if (!defaults) {
    lookup.docDefaultsItalic = false;
    return false;
  }
  const rPrDefault = directChild(defaults, 'w:rPrDefault');
  const rPr = directChild(rPrDefault, 'w:rPr');
  lookup.docDefaultsItalic = directItalic(rPr) === true;
  return lookup.docDefaultsItalic;
}

const PART_ROOT_TAGS = new Set(['w:body', 'w:hdr', 'w:ftr', 'w:footnote', 'w:endnote', 'w:comment', 'w:document']);

/** The `w:tblStyle` id of the enclosing table that holds a run, if inside a table. */
function tableStyleIdOf(run: Element): string | undefined {
  let node: Node | null = run.parentNode;
  while (node) {
    const el = node as Element;
    if (el.nodeType === 1) {
      if (el.tagName === 'w:tbl') {
        const tblPr = directChild(el, 'w:tblPr');
        return directChild(tblPr, 'w:tblStyle')?.getAttribute('w:val') ?? undefined;
      }
      if (PART_ROOT_TAGS.has(el.tagName)) return undefined;
    }
    node = node.parentNode;
  }
  return undefined;
}

/** The `w:pStyle` id of the paragraph that holds a run. */
function paragraphStyleIdOf(run: Element): string | undefined {
  let node: Node | null = run.parentNode;
  while (node && (node as Element).tagName !== 'w:p') node = node.parentNode;
  const pPr = directChild(node as Element | null, 'w:pPr');
  return directChild(pPr, 'w:pStyle')?.getAttribute('w:val') ?? undefined;
}

/** Whether a run shows italic: its own formatting, else the toggle of docDefaults, table style, paragraph style and character style. */
export function effectiveRunItalic(run: Element, ast?: DocumentAST): boolean {
  const rPr = run.getElementsByTagName('w:rPr')[0];
  const direct = directItalic(rPr);
  if (direct !== undefined) return direct;
  const lookup = ast ? lookups.get(ast) : undefined;
  if (!lookup) return false;
  const fromDocDefaults = italicOfDocDefaults(lookup);
  const fromTable = italicOfStyle(lookup, tableStyleIdOf(run));
  const fromParagraph = italicOfStyle(lookup, paragraphStyleIdOf(run));
  const fromCharacter = italicOfStyle(lookup, rPr ? directChild(rPr, 'w:rStyle')?.getAttribute('w:val') : undefined);
  const count = (fromDocDefaults ? 1 : 0) + (fromTable ? 1 : 0) + (fromParagraph ? 1 : 0) + (fromCharacter ? 1 : 0);
  return count % 2 === 1;
}
