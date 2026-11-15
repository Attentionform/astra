/**
 * Text and line properties inside chart parts (DrawingML `a:` elements), shared by the
 * title format, the chart-wide font, axis label settings and gridline colours (0.5.3).
 * Everything is text-level so the rest of the part stays byte-identical.
 */
import { findDirectChild, setAttribute, tagEnd, tagNameOf } from './chart-xml-spans';

export interface RunProps {
  /** Size in hundredths of a point (`sz`). */
  sz?: number;
  bold?: boolean;
  /** Six hex digits, no '#'. */
  color?: string;
  /** Latin typeface. */
  family?: string;
}

const FILL_TAGS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
// Children that must come after the fill in CT_TextCharacterProperties.
const AFTER_FILL = [
  'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill',
  'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst',
];
const AFTER_LATIN = ['a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'];

export function expand(elementXml: string): string {
  const gt = tagEnd(elementXml, 0);
  if (gt < 0 || elementXml[gt - 1] !== '/') return elementXml;
  return `${elementXml.slice(0, gt - 1)}>${'</' + tagNameOf(elementXml)}>`;
}

export function insertBefore(elementXml: string, laterTags: readonly string[], child: string): string {
  for (const tag of laterTags) {
    const span = findDirectChild(elementXml, tag);
    if (span) return elementXml.slice(0, span.start) + child + elementXml.slice(span.start);
  }
  const close = elementXml.lastIndexOf('</');
  return elementXml.slice(0, close) + child + elementXml.slice(close);
}

/** Replaces the direct fill child of a DrawingML properties element, or inserts one before the first of `laterTags`. */
export function setFillChild(elementXml: string, fillXml: string, laterTags: readonly string[]): string {
  const out = expand(elementXml);
  const existing = FILL_TAGS.map((tag) => findDirectChild(out, tag)).find((span) => span !== null);
  return existing
    ? out.slice(0, existing.start) + fillXml + out.slice(existing.end)
    : insertBefore(out, laterTags, fillXml);
}

export function solidFillXml(color: string): string {
  return `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
}

/** Sets the given properties on one `a:rPr`, `a:defRPr` or `a:endParaRPr` element; omitted properties stay. */
export function applyRunProps(elementXml: string, props: RunProps): string {
  let out = elementXml;
  if (props.sz !== undefined) out = setAttribute(out, 'sz', String(props.sz));
  if (props.bold !== undefined) out = setAttribute(out, 'b', props.bold ? '1' : '0');
  if (props.color === undefined && props.family === undefined) return out;
  out = expand(out);
  if (props.color !== undefined) {
    const existing = FILL_TAGS.map((tag) => findDirectChild(out, tag)).find((span) => span !== null);
    out = existing
      ? out.slice(0, existing.start) + solidFillXml(props.color) + out.slice(existing.end)
      : insertBefore(out, AFTER_FILL, solidFillXml(props.color));
  }
  if (props.family !== undefined) {
    const latin = `<a:latin typeface="${props.family.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"/>`;
    const existing = findDirectChild(out, 'a:latin');
    out = existing
      ? out.slice(0, existing.start) + latin + out.slice(existing.end)
      : insertBefore(out, AFTER_LATIN, latin);
  }
  return out;
}

/** Applies run properties to every `a:rPr` and `a:defRPr` inside a block (a title, a text-properties block). */
export function applyRunPropsDeep(blockXml: string, props: RunProps): string {
  const re = /<a:(?:rPr|defRPr)(?=[\s/>])/g;
  const starts: number[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(blockXml)) !== null) starts.push(m.index);
  let out = blockXml;
  for (const start of starts.reverse()) {
    const name = out.startsWith('<a:rPr', start) ? 'a:rPr' : 'a:defRPr';
    const gt = tagEnd(out, start);
    let end: number;
    if (gt >= 0 && out[gt - 1] === '/') {
      end = gt + 1;
    } else {
      const close = out.indexOf(`</${name}>`, start);
      if (close < 0) continue;
      end = close + `</${name}>`.length;
    }
    out = out.slice(0, start) + applyRunProps(out.slice(start, end), props) + out.slice(end);
  }
  return out;
}

/** `c:txPr` holding one default run-properties element, for elements that carry no rich text of their own. */
export function textPropertiesXml(qualify: (local: string) => string, props: RunProps, rotationDeg?: number): string {
  const defRPr = applyRunProps('<a:defRPr/>', props);
  const rot = rotationDeg === undefined ? '' : ` rot="${Math.round(rotationDeg * 60000)}" vert="horz"`;
  return `<${qualify('txPr')}><a:bodyPr${rot}/><a:lstStyle/><a:p><a:pPr>${defRPr}</a:pPr><a:endParaRPr lang="en-US"/></a:p></${qualify('txPr')}>`;
}
