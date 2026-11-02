/**
 * Italic of a run for the HTML and Markdown exports, from the run, its character style and its paragraph
 * style. A Subtitle or Quote paragraph is italic through its style, with no `w:i` on any run, so reading
 * run properties alone exported it as plain text.
 *
 * Italic is a toggle property (ECMA-376 17.7.3): direct formatting on the run (`w:i` on or off) decides
 * alone. Without it, the paragraph style and the character style toggle each other, so a paragraph style
 * that is italic and a character style that is italic give plain text. Within one style, the closest
 * definition in the `w:basedOn` chain wins.
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
}

const lookups = new WeakMap<DocumentAST, StyleLookup>();

/** Set (or, without a parser, clear) the styles that `effectiveRunItalic` reads for this AST. Call at the start of an export. */
export function useStylesForExport(ast: DocumentAST, parser: unknown): void {
  const source = parser as { getStyleElements?: () => Map<string, Element> } | null | undefined;
  if (source && typeof source.getStyleElements === 'function') {
    try {
      lookups.set(ast, { elements: source.getStyleElements(), italic: new Map() });
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

/** The `w:pStyle` id of the paragraph that holds a run. */
function paragraphStyleIdOf(run: Element): string | undefined {
  let node: Node | null = run.parentNode;
  while (node && (node as Element).tagName !== 'w:p') node = node.parentNode;
  const pPr = directChild(node as Element | null, 'w:pPr');
  return directChild(pPr, 'w:pStyle')?.getAttribute('w:val') ?? undefined;
}

/** Whether a run shows italic: its own formatting, else its paragraph style against its character style. */
export function effectiveRunItalic(run: Element, ast?: DocumentAST): boolean {
  const rPr = run.getElementsByTagName('w:rPr')[0];
  const direct = directItalic(rPr);
  if (direct !== undefined) return direct;
  const lookup = ast ? lookups.get(ast) : undefined;
  if (!lookup) return false;
  const fromParagraph = italicOfStyle(lookup, paragraphStyleIdOf(run));
  const fromCharacter = italicOfStyle(lookup, rPr ? directChild(rPr, 'w:rStyle')?.getAttribute('w:val') : undefined);
  return fromParagraph !== fromCharacter;
}
