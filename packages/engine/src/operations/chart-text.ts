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

const DRAWINGML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The prefix a part binds to the DrawingML main namespace ('' for the default namespace); undefined when it declares none. */
export function drawingPrefixOf(xml: string): string | undefined {
  const m = new RegExp(`xmlns(?::([\\w.-]+))?\\s*=\\s*["']${DRAWINGML_NS.replace(/[.]/g, '\\.')}["']`).exec(xml);
  return m ? (m[1] ?? '') : undefined;
}

/** A DrawingML name in the given prefix: `q('d', 'rPr')` is `d:rPr`, `q('', 'rPr')` is `rPr`. */
const q = (prefix: string, local: string): string => (prefix ? `${prefix}:${local}` : local);

const FILL_LOCALS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
// Children that must come after the fill in CT_TextCharacterProperties.
const AFTER_FILL_LOCALS = [
  'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill',
  'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst',
];
const AFTER_LATIN_LOCALS = ['ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
const withPrefix = (prefix: string, locals: readonly string[]): string[] => locals.map((l) => q(prefix, l));
/** The prefix of an element's own tag name. */
const prefixOfElement = (elementXml: string): string => {
  const name = tagNameOf(elementXml);
  const colon = name.indexOf(':');
  return colon > 0 ? name.slice(0, colon) : '';
};

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
export function setFillChild(elementXml: string, fillXml: string, laterTags: readonly string[], prefix = 'a'): string {
  const out = expand(elementXml);
  const existing = withPrefix(prefix, FILL_LOCALS).map((tag) => findDirectChild(out, tag)).find((span) => span !== null);
  return existing
    ? out.slice(0, existing.start) + fillXml + out.slice(existing.end)
    : insertBefore(out, laterTags, fillXml);
}

export function solidFillXml(color: string, prefix = 'a'): string {
  return `<${q(prefix, 'solidFill')}><${q(prefix, 'srgbClr')} val="${color}"/></${q(prefix, 'solidFill')}>`;
}

/** Sets the given properties on one `a:rPr`, `a:defRPr` or `a:endParaRPr` element; omitted properties stay. */
export function applyRunProps(elementXml: string, props: RunProps): string {
  const prefix = prefixOfElement(elementXml);
  let out = elementXml;
  if (props.sz !== undefined) out = setAttribute(out, 'sz', String(props.sz));
  if (props.bold !== undefined) out = setAttribute(out, 'b', props.bold ? '1' : '0');
  if (props.color === undefined && props.family === undefined) return out;
  out = expand(out);
  if (props.color !== undefined) {
    const existing = withPrefix(prefix, FILL_LOCALS).map((tag) => findDirectChild(out, tag)).find((span) => span !== null);
    out = existing
      ? out.slice(0, existing.start) + solidFillXml(props.color, prefix) + out.slice(existing.end)
      : insertBefore(out, withPrefix(prefix, AFTER_FILL_LOCALS), solidFillXml(props.color, prefix));
  }
  if (props.family !== undefined) {
    const latin = `<${q(prefix, 'latin')} typeface="${props.family.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"/>`;
    const existing = findDirectChild(out, q(prefix, 'latin'));
    out = existing
      ? out.slice(0, existing.start) + latin + out.slice(existing.end)
      : insertBefore(out, withPrefix(prefix, AFTER_LATIN_LOCALS), latin);
  }
  return out;
}

/** Applies run properties to every `rPr` and `defRPr` of the DrawingML prefix inside a block (a title, a text-properties block). */
export function applyRunPropsDeep(blockXml: string, props: RunProps, prefix = 'a'): string {
  const escaped = prefix.replace(/[.]/g, '\\.');
  const re = new RegExp(`<${prefix ? `${escaped}:` : ''}(?:rPr|defRPr)(?=[\\s/>])`, 'g');
  const starts: number[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(blockXml)) !== null) starts.push(m.index);
  let out = blockXml;
  for (const start of starts.reverse()) {
    const name = out.startsWith(`<${q(prefix, 'rPr')}`, start) && /[\s/>]/.test(out[start + q(prefix, 'rPr').length + 1] ?? '') ? q(prefix, 'rPr') : q(prefix, 'defRPr');
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
export function textPropertiesXml(qualify: (local: string) => string, props: RunProps, rotationDeg?: number, prefix = 'a'): string {
  const defRPr = applyRunProps(`<${q(prefix, 'defRPr')}/>`, props);
  const rot = rotationDeg === undefined ? '' : ` rot="${Math.round(rotationDeg * 60000)}" vert="horz"`;
  return `<${qualify('txPr')}><${q(prefix, 'bodyPr')}${rot}/><${q(prefix, 'lstStyle')}/><${q(prefix, 'p')}><${q(prefix, 'pPr')}>${defRPr}</${q(prefix, 'pPr')}><${q(prefix, 'endParaRPr')} lang="en-US"/></${q(prefix, 'p')}></${qualify('txPr')}>`;
}
