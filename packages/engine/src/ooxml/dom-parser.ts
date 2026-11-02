/**
 * The XML parser the engine uses for every package part.
 *
 * Word writes XML 1.0. In XML 1.0 only a line feed, a carriage return and a carriage return followed by
 * a line feed are line ends. U+2028 (line separator) and U+0085 (next line) are ordinary characters that
 * can sit in the text of a run. The parser library applies the XML 1.1 rule and turns both into a line
 * feed, so a document that held either character came out of a save with different text.
 *
 * Both characters are written as character references before the text reaches the library. A reference is
 * kept as the character itself, and the serializer writes the character back, so the part keeps its bytes.
 * Comments, CDATA sections and processing instructions are left alone: a reference there would be literal text.
 */
import { DOMParser as XmldomDOMParser, XMLSerializer } from '@xmldom/xmldom';

// Built from code points so that no editor or tool can turn the characters into line breaks in this source.
const NEXT_LINE = String.fromCharCode(0x85);
const LINE_SEPARATOR = String.fromCharCode(0x2028);
const SEPARATORS = new RegExp(`[${NEXT_LINE}${LINE_SEPARATOR}]`);
const NEXT_LINE_ALL = new RegExp(NEXT_LINE, 'g');
const LINE_SEPARATOR_ALL = new RegExp(LINE_SEPARATOR, 'g');
const PROTECTED_SEGMENTS = /(<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>)/;

/** Writes U+0085 and U+2028 as character references, outside comments, CDATA sections and processing instructions. */
export function protectLineSeparators(xml: string): string {
  if (!SEPARATORS.test(xml)) return xml;
  return xml
    .split(PROTECTED_SEGMENTS)
    .map((segment, index) =>
      index % 2 === 1 ? segment : segment.replace(NEXT_LINE_ALL, '&#x85;').replace(LINE_SEPARATOR_ALL, '&#x2028;')
    )
    .join('');
}

export class DOMParser extends XmldomDOMParser {
  parseFromString(source: string, mimeType?: string): Document {
    return super.parseFromString(protectLineSeparators(source), mimeType);
  }
}

export { XMLSerializer };
