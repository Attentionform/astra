import { DOMParser } from './dom-parser';

/**
 * The `Quote` paragraph style: an indented block with a left rule and italic
 * text in the body font. Every built-in template carries it, and compose adds
 * it to an opened document that lacks one, so a Markdown `>` quote is a real
 * Word style the author can restyle in one step. Nested quote levels deeper
 * than the first add a direct indent on the paragraph.
 */
export const QUOTE_STYLE_ID = 'Quote';

/** Left indent of a level-1 quote paragraph, in twips (0.5 inch). Each nested level adds the same amount. */
export const QUOTE_INDENT_TWIPS = 720;

/** The style element as it sits inside an existing styles part (the part declares the `w` prefix). */
export const QUOTE_STYLE_BODY = `<w:style w:type="paragraph" w:styleId="Quote">
    <w:name w:val="Quote"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="29"/>
    <w:qFormat/>
    <w:pPr>
      <w:pBdr>
        <w:left w:val="single" w:sz="18" w:space="10" w:color="94A3B8"/>
      </w:pBdr>
      <w:spacing w:before="120" w:after="120"/>
      <w:ind w:left="${QUOTE_INDENT_TWIPS}"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
      <w:color w:val="475569"/>
    </w:rPr>
  </w:style>`;

const QUOTE_STYLE_XML = QUOTE_STYLE_BODY.replace('<w:style ', '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ');

/**
 * Adds the `Quote` style to a styles part that lacks it. Returns true when the
 * style was added, false when the document already declares a `Quote` style (the
 * author's own definition stays) or has no styles part.
 */
export function ensureQuoteStyle(stylesDom: Document | null): boolean {
  if (!stylesDom) return false;
  const root = stylesDom.getElementsByTagName('w:styles')[0];
  if (!root) return false;
  for (const style of Array.from(stylesDom.getElementsByTagName('w:style'))) {
    if (style.getAttribute('w:styleId') === QUOTE_STYLE_ID) return false;
  }
  const fragment = new DOMParser().parseFromString(QUOTE_STYLE_XML, 'text/xml').documentElement;
  if (!fragment) return false;
  root.appendChild(stylesDom.importNode(fragment, true));
  return true;
}
