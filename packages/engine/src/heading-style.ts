import type { DocumentAST, Node } from './ast';

/**
 * Which paragraph styles are headings, and at what level. One rule for the structural summary
 * (`isHeading`, `headingLevel`) and for the `outline` tier, so the two cannot disagree.
 */

/**
 * The outline level of a style id that already counted as a heading (its lowercase id contains `heading`, `title` or
 * `subtitle`). `Heading1` to `Heading9` (spaces ignored) give 1 to 9, `Title` gives 1 and `Subtitle` gives 2. Any other
 * matching style (`TOCHeading`, `Heading1Char`, `BookTitle`) keeps level 1. Which paragraphs are headings is decided by
 * the caller and is not changed here.
 */
export function headingLevelOfStyle(lowerStyleId: string): number {
  const id = lowerStyleId.replace(/\s+/g, '');
  const numbered = /^heading([1-9])$/.exec(id);
  if (numbered) return Number(numbered[1]);
  if (id === 'subtitle') return 2;
  return 1;
}

/** The level of a heading style id, or null when the style is not a heading style (any case; `h1` is not one). */
export function headingLevelOfStyleId(styleId: string | undefined): number | null {
  if (!styleId) return null;
  const lower = styleId.toLowerCase();
  if (!(lower.includes('heading') || lower.includes('title') || lower.includes('subtitle'))) return null;
  return headingLevelOfStyle(lower);
}

/**
 * The style id written in a paragraph's own `w:pPr/w:pStyle`, or undefined. Only the paragraph's own properties
 * count: a table, or a paragraph that holds a textbox, does not take the style of a paragraph nested inside it.
 */
export function ownParagraphStyleId(paragraph: Element): string | undefined {
  if (paragraph.tagName !== 'w:p') return undefined;
  for (let i = 0; i < paragraph.childNodes.length; i++) {
    const pPr = paragraph.childNodes[i] as Element;
    if (pPr.nodeType !== 1 || pPr.tagName !== 'w:pPr') continue;
    for (let k = 0; k < pPr.childNodes.length; k++) {
      const el = pPr.childNodes[k] as Element;
      if (el.nodeType === 1 && el.tagName === 'w:pStyle') return el.getAttribute('w:val') || undefined;
    }
    return undefined;
  }
  return undefined;
}

/**
 * The style and level of a body node that is a heading, or null: a paragraph whose own style is a heading style.
 * The test is on the XML (a `w:p`), not on the AST node type: the parser types a paragraph that holds a display
 * equation, a chart or an image as `equation`, `chart` or `image`, and the structural summary still calls it a heading.
 */
export function headingOfNode(node: Pick<Node, 'xml'>): { style: string; level: number } | null {
  const style = ownParagraphStyleId(node.xml);
  const level = headingLevelOfStyleId(style);
  return style !== undefined && level !== null ? { style, level } : null;
}

/** Headings in the whole document, by the rule of the `outline` tier. */
export function countHeadings(ast: Pick<DocumentAST, 'sections'>): number {
  let n = 0;
  for (const sec of ast.sections) for (const node of sec.nodes) if (headingOfNode(node)) n += 1;
  return n;
}
