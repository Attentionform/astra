/**
 * Network references in a .docx, and how a render treats them.
 *
 * A .docx can point outside itself: a picture linked from a URL, an attached template on a file share, a
 * linked OLE object, a field that includes text from another file. Word and LibreOffice may fetch such a
 * target when they open the file. A `\\host\share\x` target makes Windows send the user's login hash to
 * `host`, and a dead `http://` host stalls the render.
 *
 * By default a render works on a temporary copy in which these references are retargeted to an inert local
 * name (`neutraliseReferences`); the user's file is never changed. With `allowNetworkReferences` the copy is
 * left alone and a read-only scan only reports what it recognised.
 *
 * What is covered, exactly (nothing else is inspected, so no claim of "no network traffic" is made):
 *   - every relationship (in every `.rels` part) other than a hyperlink whose target is not a plain local
 *     path: retargeted in place, so relationship ids stay valid;
 *   - the instruction of fields that load external content (INCLUDEPICTURE, INCLUDETEXT, LINK, IMPORT) when
 *     the file argument is not a plain local path, including instructions split across several runs and
 *     `w:fldSimple`; DDE and DDEAUTO fields are always neutralised (they can start programs).
 * Every part is looked at by content, not by name or extension (a renderer reads a part according to its
 * declared content type, so a main document part could be hidden under any name), except media-looking parts
 * over 16 MB, which are skipped.
 * Not inspected: VML `o:href` attributes, mail-merge data-source connection strings, DATABASE fields, content
 * inside embedded objects and imported chunks (altChunk), and anything a renderer does that is not driven by
 * the structures above.
 *
 * Diagnostics are minimal on purpose: scheme class, relationship or field type and part name only. A target
 * itself (host, path, query, credentials) is never put into a result, a log or an error.
 */

import { isEncryptedOoxmlPackage } from './cfb';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import AdmZip from '../adm-zip-safe';
import { checkZipBombCeilings } from '../zip-hardening';
import { codedError } from './renderer-preference';

export const NETWORK_REFERENCES_ENV = 'ASTRA_ALLOW_NETWORK_REFERENCES';
export const INERT_TARGET = 'astra-reference-not-loaded';
/** Report lists are capped so a document with thousands of links cannot produce a huge payload; counts stay exact. */
export const REFERENCE_LIST_CAP = 50;

export type ReferenceClass = 'network' | 'unknown-scheme' | 'unparseable' | 'dde' | 'local' | 'embedded-content';
export type ReferenceScheme = 'http' | 'https' | 'ftp' | 'file' | 'unc' | 'other';

export interface ReferenceEntry {
  /** Package part holding the reference (document-internal name, e.g. `word/_rels/document.xml.rels`). */
  part: string;
  /** Relationship id, or `field#N` for the Nth recognised field in the part. */
  id: string;
  kind: 'relationship' | 'field';
  /** Relationship type's last segment (`image`, `attachedTemplate`) or the field type (`INCLUDEPICTURE`). */
  type: string;
  class: ReferenceClass;
  scheme?: ReferenceScheme;
}

/** Parse the option from a flag, parameter or ASTRA_ALLOW_NETWORK_REFERENCES. Empty means "not set". */
export function parseAllowNetworkReferences(
  raw: unknown,
  source: 'flag' | 'parameter' | typeof NETWORK_REFERENCES_ENV = 'parameter'
): boolean | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (typeof raw === 'boolean') return raw;
  if (typeof raw === 'string') {
    const v = raw.trim().toLowerCase();
    if (v === '') return undefined;
    if (['1', 'true', 'yes', 'on'].includes(v)) return true;
    if (['0', 'false', 'no', 'off'].includes(v)) return false;
  }
  throw codedError(
    'E_VALIDATION',
    `E_VALIDATION: ${source} allowNetworkReferences must be true or false (1/0, yes/no, on/off).`,
    `Use ${NETWORK_REFERENCES_ENV}=1 to allow, ${NETWORK_REFERENCES_ENV}=0 (or unset it) to neutralise network references.`
  );
}

/** Explicit per-call value (including an explicit `false`) > ASTRA_ALLOW_NETWORK_REFERENCES > false. */
export function effectiveAllowNetworkReferences(
  explicit: boolean | string | undefined,
  env: NodeJS.ProcessEnv = process.env
): boolean {
  const fromCall = parseAllowNetworkReferences(explicit, 'parameter');
  if (fromCall !== undefined) return fromCall;
  return parseAllowNetworkReferences(env[NETWORK_REFERENCES_ENV], NETWORK_REFERENCES_ENV) ?? false;
}

// ---------------------------------------------------------------------------------------------------------
// Classification of one target (a relationship Target or a field's file argument)
// ---------------------------------------------------------------------------------------------------------

export interface TargetClassification {
  class: 'network' | 'unknown-scheme' | 'unparseable' | 'local';
  scheme?: ReferenceScheme;
}

// eslint-disable-next-line no-control-regex -- the point is to detect control characters in a target
const CONTROL_CHARS = /[\u0000-\u001f\u007f]/;
const MAX_TARGET_LENGTH = 4096;
const KNOWN_NETWORK_SCHEMES: Record<string, ReferenceScheme> = {
  http: 'http',
  https: 'https',
  ftp: 'ftp',
  ftps: 'ftp',
};

function classifyFileUri(t: string): TargetClassification {
  let r = t.slice('file:'.length).replace(/\\/g, '/');
  if (r.startsWith('//')) {
    const after = r.slice(2);
    const slash = after.indexOf('/');
    const authority = slash < 0 ? after : after.slice(0, slash);
    const rest = slash < 0 ? '' : after.slice(slash);
    // `file://host/share` is a network location; `file:///C:/x` (empty authority), `file://localhost/` and a
    // drive written as the authority (`file://C:/x`) are local.
    if (authority !== '' && authority.toLowerCase() !== 'localhost' && !/^[a-z][:|]$/i.test(authority)) {
      return { class: 'network', scheme: 'file' };
    }
    r = rest;
  }
  // `file:////host/share` and `file:///\\host\share` carry a UNC path after the empty authority.
  if (r.startsWith('//')) return { class: 'network', scheme: 'file' };
  return { class: 'local' };
}

function classifyOne(raw: string): TargetClassification {
  const t = raw.trim();
  if (t === '') return { class: 'local' };
  if (t.length > MAX_TARGET_LENGTH || CONTROL_CHARS.test(t)) return { class: 'unparseable' };
  if (/^[\\/]{2}/.test(t)) {
    // `\\?\C:\x` and `\\.\C:\x` are local device paths; `\\?\UNC\host\x` is a network path; any other
    // device namespace is not understood, so it is not trusted.
    if (/^[\\/]{2}[?.][\\/]/.test(t)) {
      const rest = t.slice(4);
      if (/^[a-z]:([\\/]|$)/i.test(rest)) return { class: 'local' };
      if (/^unc[\\/]/i.test(rest)) return { class: 'network', scheme: 'unc' };
      return { class: 'unparseable' };
    }
    return { class: 'network', scheme: 'unc' };
  }
  const m = /^([a-z][a-z0-9+.-]*):/i.exec(t);
  if (m) {
    const scheme = m[1].toLowerCase();
    if (scheme.length === 1) return { class: 'local' }; // a drive letter
    if (scheme === 'file') return classifyFileUri(t);
    const known = KNOWN_NETWORK_SCHEMES[scheme];
    if (known) return { class: 'network', scheme: known };
    return { class: 'unknown-scheme', scheme: 'other' };
  }
  // A scheme separator that the pattern above did not recognise (odd leading characters) is not trusted.
  if (/:[\\/]{2}/.test(t)) return { class: 'unparseable' };
  return { class: 'local' };
}

const SEVERITY: Record<TargetClassification['class'], number> = { network: 3, 'unknown-scheme': 2, unparseable: 1, local: 0 };

/**
 * Classify a target. The raw text and up to three rounds of percent-decoding are all checked and the most
 * severe result wins, so `%5C%5Chost%5Cshare` cannot hide a UNC path.
 */
export function classifyTarget(raw: string): TargetClassification {
  const variants: string[] = [raw];
  let cur = raw;
  for (let i = 0; i < 3; i++) {
    let next: string;
    try {
      next = decodeURIComponent(cur);
    } catch {
      // Malformed escapes: keep the text we have and also look at a lenient decode.
      next = cur.replace(/%([0-9a-f]{2})/gi, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
    }
    if (next === cur) break;
    variants.push(next);
    cur = next;
  }
  let best: TargetClassification = { class: 'local' };
  for (const v of variants) {
    const c = classifyOne(v);
    if (SEVERITY[c.class] > SEVERITY[best.class]) best = c;
  }
  return best;
}

// ---------------------------------------------------------------------------------------------------------
// Field instruction parsing
// ---------------------------------------------------------------------------------------------------------

interface FieldToken {
  raw: string;
  /** Text with quotes removed and `\\` / `\"` unescaped. */
  value: string;
  quoted: boolean;
  start: number;
  end: number;
}

const DDE_FIELDS = new Set(['DDE', 'DDEAUTO']);
/** Field type -> which non-switch argument names the external file (0-based). */
const FILE_FIELDS: Record<string, number> = { INCLUDEPICTURE: 0, INCLUDETEXT: 0, IMPORT: 0, LINK: 1 };
const MAX_INSTRUCTION_LENGTH = 32768;

/** Split an instruction (after the field type) into quoted and unquoted tokens, keeping their source spans. */
function tokenizeInstruction(text: string, from: number): FieldToken[] {
  const tokens: FieldToken[] = [];
  let i = from;
  while (i < text.length) {
    if (/\s/.test(text[i])) {
      i++;
      continue;
    }
    const start = i;
    if (text[i] === '"') {
      i++;
      let value = '';
      while (i < text.length && text[i] !== '"') {
        if (text[i] === '\\' && i + 1 < text.length && (text[i + 1] === '\\' || text[i + 1] === '"')) {
          value += text[i + 1];
          i += 2;
        } else {
          value += text[i];
          i++;
        }
      }
      if (i < text.length) i++; // closing quote (an unterminated quote runs to the end)
      tokens.push({ raw: text.slice(start, i), value, quoted: true, start, end: i });
    } else {
      while (i < text.length && !/\s/.test(text[i])) i++;
      const raw = text.slice(start, i);
      tokens.push({ raw, value: raw, quoted: false, start, end: i });
    }
  }
  return tokens;
}

function isSwitch(t: FieldToken): boolean {
  return !t.quoted && /^\\[^\\]/.test(t.raw);
}

interface FieldVerdict {
  type: string;
  /** New instruction text when the field must be neutralised. */
  rewritten?: string;
  class: ReferenceClass;
  scheme?: ReferenceScheme;
}

/** Decide what a field instruction needs. Returns null when it is not a field type of interest. */
export function inspectFieldInstruction(instruction: string): FieldVerdict | null {
  const m = /^\s*([A-Za-z]+)/.exec(instruction);
  if (!m) return null;
  const type = m[1].toUpperCase();
  const isDde = DDE_FIELDS.has(type);
  const fileIndex = FILE_FIELDS[type];
  if (!isDde && fileIndex === undefined) return null;
  if (isDde) return { type, rewritten: ' QUOTE "" ', class: 'dde' };
  if (instruction.length > MAX_INSTRUCTION_LENGTH) return { type, rewritten: ' QUOTE "" ', class: 'unparseable' };
  const args = tokenizeInstruction(instruction, m[0].length).filter((t) => !isSwitch(t));
  const target = args[fileIndex];
  if (!target) return { type, class: 'local' };
  // An unquoted token is classified both as written and with doubled backslashes collapsed.
  const classes = [classifyTarget(target.value)];
  if (!target.quoted) classes.push(classifyTarget(target.raw.replace(/\\\\/g, '\\')));
  let best = classes[0];
  for (const c of classes) if (SEVERITY[c.class] > SEVERITY[best.class]) best = c;
  if (best.class === 'local') return { type, class: 'local' };
  const rewritten = `${instruction.slice(0, target.start)}"${INERT_TARGET}"${instruction.slice(target.end)}`;
  return { type, rewritten, class: best.class, scheme: best.scheme };
}

// ---------------------------------------------------------------------------------------------------------
// Package scan / rewrite
// ---------------------------------------------------------------------------------------------------------

const W_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
/** Parts bigger than this that carry a media-like name are not read (inflating them to look inside is not worth it). */
const SNIFF_LIMIT_BYTES = 16 * 1024 * 1024;
const MEDIA_NAME = /\.(?:png|jpe?g|gif|bmp|tiff?|emf|wmf|wdp|svg|ico|mp3|mp4|m4a|wav|avi|mov|mpe?g|wmv|webm|ogg|ttf|otf|odttf|fntdata|bin|pdf|zip|docx|xlsx|pptx)$/i;
const MAX_PART_BYTES = 128 * 1024 * 1024;

export interface ReferenceAnalysis {
  /** The bytes to render: the input itself when nothing had to change. */
  docx: Buffer;
  changed: boolean;
  /** References that were (or, in scan-only mode, would be) neutralised. */
  found: ReferenceEntry[];
  /** Local-path references that were kept. */
  local: ReferenceEntry[];
  /** Embedded or imported content that is not inspected. */
  notInspected: ReferenceEntry[];
}

function failClosed(message: string): Error {
  return codedError(
    'E_SECURITY',
    `E_SECURITY: references in the document could not be checked: ${message}`,
    'Repair the document, or render it with --allow-network-references only if you trust it (the renderer may then contact network locations the document names).'
  );
}

function getAttr(el: Element, local: string): string | null {
  const direct = el.getAttribute(local);
  if (direct !== null && direct !== '') return direct;
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a && (a.localName === local || a.name.endsWith(`:${local}`))) return a.value;
  }
  return null;
}

function relTypeTail(type: string): string {
  const tail = type.replace(/[#?].*$/, '').split('/').pop() ?? type;
  return tail.slice(0, 48) || 'unknown';
}

function parseXml(name: string, xml: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw failClosed(`part "${name}" carries a DOCTYPE or entity declaration`);
  const problems: string[] = [];
  const doc = new DOMParser({
    errorHandler: {
      warning: (m: string) => problems.push(String(m)),
      error: (m: string) => problems.push(String(m)),
      fatalError: (m: string) => problems.push(String(m)),
    },
  }).parseFromString(xml, 'text/xml');
  if (problems.length > 0 || !doc || !doc.documentElement) throw failClosed(`part "${name}" is not well-formed XML`);
  return doc as unknown as Document;
}

/** How a part's bytes map to text, kept so a rewritten part goes back out in the encoding it came in. */
interface PartEncoding {
  utf16?: 'le' | 'be';
  /** The bytes started with a byte order mark (UTF-8 BOM or UTF-16 BOM). */
  bom: boolean;
}

/**
 * UTF-16 XML: a BOM (FF FE or FE FF) or none, then `<` as the first character, which shows as `3C 00` (LE)
 * or `00 3C` (BE). A binary part that merely starts with FF FE is not text and is left alone.
 */
function utf16Endianness(data: Buffer): 'le' | 'be' | undefined {
  if (data.length < 4) return undefined;
  const [a, b, c, d] = [data[0], data[1], data[2], data[3]];
  if (a === 0xff && b === 0xfe && c === 0x3c && d === 0x00) return 'le';
  if (a === 0xfe && b === 0xff && c === 0x00 && d === 0x3c) return 'be';
  if (a === 0x3c && b === 0x00 && d === 0x00) return 'le';
  if (a === 0x00 && b === 0x3c && c === 0x00) return 'be';
  return undefined;
}

/** Does the part start like an XML document (after a UTF-8 BOM and white space)? */
function looksLikeXml(data: Buffer): boolean {
  let i = 0;
  if (data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf) i = 3;
  while (i < data.length && i < 64 && (data[i] === 0x20 || data[i] === 0x09 || data[i] === 0x0a || data[i] === 0x0d)) i++;
  return data[i] === 0x3c;
}

function decodePart(name: string, data: Buffer): { text: string; enc: PartEncoding } {
  if (data.length > MAX_PART_BYTES) throw failClosed(`part "${name}" is larger than ${MAX_PART_BYTES >> 20} MB`);
  const utf16 = utf16Endianness(data);
  if (utf16) {
    if (data.length % 2 !== 0) throw failClosed(`part "${name}" is not valid UTF-16`);
    let bytes = data;
    if (utf16 === 'be') {
      bytes = Buffer.from(data); // swap on a copy, the zip entry stays as it was
      bytes.swap16();
    }
    let text = bytes.toString('utf16le');
    const bom = text.charCodeAt(0) === 0xfeff;
    if (bom) text = text.slice(1);
    return { text, enc: { utf16, bom } };
  }
  let text = data.toString('utf8');
  const bom = text.charCodeAt(0) === 0xfeff;
  if (bom) text = text.slice(1);
  return { text, enc: { bom } };
}

function encodePart(xml: string, enc: PartEncoding): Buffer {
  if (!enc.utf16) return Buffer.from(xml, 'utf8');
  const bytes = Buffer.from((enc.bom ? '\ufeff' : '') + xml, 'utf16le');
  if (enc.utf16 === 'be') bytes.swap16();
  return bytes;
}

function serialise(doc: Document): string {
  return new XMLSerializer().serializeToString(doc as never);
}

function processRels(name: string, data: Buffer, mutate: boolean, out: ReferenceAnalysis): Buffer | null {
  const { text, enc } = decodePart(name, data);
  const doc = parseXml(name, text);
  const rels = (doc as unknown as { getElementsByTagNameNS(ns: string, ln: string): ArrayLike<Element> }).getElementsByTagNameNS('*', 'Relationship');
  let changed = false;
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i];
    const type = rel.getAttribute('Type') ?? '';
    const target = rel.getAttribute('Target') ?? '';
    const id = (rel.getAttribute('Id') ?? `#${i}`).slice(0, 64);
    const tail = relTypeTail(type);
    if (/\/hyperlink$/i.test(type.replace(/[#?].*$/, ''))) continue; // followed on click, not on open
    const external = /^external$/i.test(rel.getAttribute('TargetMode') ?? '');
    const cls = classifyTarget(target);
    if (cls.class === 'local') {
      if (external) {
        out.local.push({ part: name, id, kind: 'relationship', type: tail, class: 'local' });
      } else if (/(?:aFChunk|oleObject|package)$/i.test(tail)) {
        out.notInspected.push({ part: name, id, kind: 'relationship', type: tail, class: 'embedded-content' });
      }
      continue;
    }
    out.found.push({ part: name, id, kind: 'relationship', type: tail, class: cls.class, ...(cls.scheme ? { scheme: cls.scheme } : {}) });
    if (mutate) {
      rel.setAttribute('Target', INERT_TARGET);
      rel.setAttribute('TargetMode', 'External');
      changed = true;
    }
  }
  return changed ? encodePart(serialise(doc), enc) : null;
}

interface FieldFrame {
  nodes: Element[];
  separated: boolean;
}

function textOf(el: Element): string {
  return el.textContent ?? '';
}

function setInstructionText(first: Element, rest: Element[], text: string): void {
  while (first.firstChild) first.removeChild(first.firstChild);
  first.appendChild(first.ownerDocument.createTextNode(text));
  first.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  for (const el of rest) {
    while (el.firstChild) el.removeChild(el.firstChild);
  }
}

function setFldSimpleInstruction(el: Element, text: string): void {
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a && (a.localName === 'instr' || a.name.endsWith(':instr'))) {
      a.value = text;
      a.nodeValue = text;
      return;
    }
  }
}

function processFieldsPart(name: string, data: Buffer, mutate: boolean, out: ReferenceAnalysis): Buffer | null {
  const utf16 = utf16Endianness(data) !== undefined;
  if (!utf16 && !looksLikeXml(data)) return null;
  // Cheap pre-filter on the raw bytes: most parts hold no field at all. `InstrText` also matches `delInstrText`.
  // A UTF-16 part is decoded first (it is rare), because its bytes never match a UTF-8 search string.
  let decoded: { text: string; enc: PartEncoding } | undefined;
  if (utf16) {
    decoded = decodePart(name, data);
    if (!decoded.text.includes('nstrText') && !decoded.text.includes('fldSimple')) return null;
  } else if (!data.includes('nstrText') && !data.includes('fldSimple')) {
    return null;
  }
  const { text, enc } = decoded ?? decodePart(name, data);
  const doc = parseXml(name, text);
  let ordinal = 0;
  let changed = false;

  const record = (v: FieldVerdict): void => {
    ordinal++;
    const entry: ReferenceEntry = { part: name, id: `field#${ordinal}`, kind: 'field', type: v.type, class: v.class, ...(v.scheme ? { scheme: v.scheme } : {}) };
    if (v.class === 'local') out.local.push(entry);
    else out.found.push(entry);
  };

  const finalizeComplex = (nodes: Element[]): void => {
    if (nodes.length === 0) return;
    const instruction = nodes.map(textOf).join('');
    const v = inspectFieldInstruction(instruction);
    if (!v) return;
    record(v);
    if (mutate && v.rewritten !== undefined) {
      setInstructionText(nodes[0], nodes.slice(1), v.rewritten);
      changed = true;
    }
  };

  const stack: FieldFrame[] = [];
  const todo: Node[] = [doc.documentElement];
  while (todo.length > 0) {
    const node = todo.pop() as Node;
    if (node.nodeType === 1) {
      const el = node as Element;
      if (el.namespaceURI !== null && W_NAMESPACES.has(el.namespaceURI)) {
        const ln = el.localName;
        if (ln === 'fldChar') {
          const kind = getAttr(el, 'fldCharType');
          if (kind === 'begin') stack.push({ nodes: [], separated: false });
          else if (kind === 'separate') {
            if (stack.length > 0) stack[stack.length - 1].separated = true;
          } else if (kind === 'end') {
            const frame = stack.pop();
            if (frame) finalizeComplex(frame.nodes);
          }
        } else if (ln === 'instrText' || ln === 'delInstrText') {
          const top = stack.length > 0 ? stack[stack.length - 1] : undefined;
          if (top && !top.separated) top.nodes.push(el);
          else finalizeComplex([el]); // an instruction outside a begin..separate span: judged on its own
        } else if (ln === 'fldSimple') {
          const instr = getAttr(el, 'instr');
          if (instr !== null) {
            const v = inspectFieldInstruction(instr);
            if (v) {
              record(v);
              if (mutate && v.rewritten !== undefined) {
                setFldSimpleInstruction(el, v.rewritten);
                changed = true;
              }
            }
          }
        }
      }
      for (let c = el.lastChild; c; c = c.previousSibling) todo.push(c);
    }
  }
  // A field that was begun and never ended is still judged (fail closed).
  while (stack.length > 0) finalizeComplex((stack.pop() as FieldFrame).nodes);
  return changed ? encodePart(serialise(doc), enc) : null;
}

/**
 * Find (and, with `mutate`, retarget) network references. Returns the original buffer object when nothing
 * had to change. Throws E_SECURITY when the package cannot be checked (fail closed).
 */
export function analyseReferences(docx: Buffer, opts: { mutate: boolean }): ReferenceAnalysis {
  const out: ReferenceAnalysis = { docx, changed: false, found: [], local: [], notInspected: [] };
  let zip: InstanceType<typeof AdmZip>;
  try {
    zip = new AdmZip(docx);
    // Only the total size bounds the work here. The per-entry ratio is a policy for opening a file; a document the
    // engine itself wrote (a blank bitmap compresses 1000:1) must still be exportable.
    checkZipBombCeilings(zip as never, { sourceCompressedBytes: docx.length, maxRatio: Number.MAX_SAFE_INTEGER });
  } catch (err: unknown) {
    if ((err as { code?: string })?.code === 'E_LIMIT' || /\[E_LIMIT\]/.test(String((err as Error)?.message))) throw err;
    // A password-protected document is not a zip package, and its content cannot be read without the password,
    // so no renderer can follow a reference in it. The renderer reports the password itself.
    if (isEncryptedOoxmlPackage(docx)) return out;
    throw failClosed('the file is not a readable zip package');
  }
  const entries = zip.getEntries();
  const seen = new Set<string>();
  for (const e of entries) {
    const key = e.entryName.replace(/\\/g, '/').toLowerCase();
    if (seen.has(key)) throw failClosed('the package contains two parts with the same name');
    seen.add(key);
  }
  const updates = new Map<string, Buffer>();
  for (const e of entries) {
    if (e.isDirectory) continue;
    const name = e.entryName.replace(/\\/g, '/');
    const lower = name.toLowerCase();
    if (lower.endsWith('.rels')) {
      const next = processRels(name, e.getData(), opts.mutate, out);
      if (next !== null) updates.set(e.entryName, next);
    } else {
      const size = e.header.size;
      if (size > SNIFF_LIMIT_BYTES && MEDIA_NAME.test(lower)) continue;
      if (size > MAX_PART_BYTES) {
        if (lower.endsWith('.xml')) throw failClosed(`part "${name}" is larger than ${MAX_PART_BYTES >> 20} MB`);
        continue;
      }
      const next = processFieldsPart(name, e.getData(), opts.mutate, out);
      if (next !== null) updates.set(e.entryName, next);
    }
  }
  if (opts.mutate && updates.size > 0) {
    for (const [entryName, bytes] of updates) zip.updateFile(entryName, bytes);
    out.docx = zip.toBuffer();
    out.changed = true;
  }
  return out;
}

/** Retarget every network reference on a copy. See the file header for the exact coverage. */
export function neutraliseReferences(docx: Buffer): ReferenceAnalysis {
  return analyseReferences(docx, { mutate: true });
}

// ---------------------------------------------------------------------------------------------------------
// What a render result says about references
// ---------------------------------------------------------------------------------------------------------

export interface ReferenceReport {
  /** The choice that applied to this render (and to a fallback engine). */
  networkReferencesAllowed: boolean;
  /** Network references were neutralised on the render copy (always false when allowed). */
  omissions: boolean;
  omittedCount: number;
  omittedReferences: ReferenceEntry[];
  /**
   * Allowed mode: references the scan recognised and left in place. The renderer MAY have contacted them;
   * ASTra does not know whether it did, and the scan is not exhaustive.
   */
  recognisedCount?: number;
  recognisedReferences?: ReferenceEntry[];
  /** Local-path links that were kept (they will not resolve on another machine). */
  localCount: number;
  /** Embedded or imported content that was not inspected. */
  notInspectedCount: number;
  /** Allowed mode only: the reporting scan failed (the render went ahead). */
  scanFailed?: boolean;
}

export function emptyReferenceReport(allowed: boolean): ReferenceReport {
  return {
    networkReferencesAllowed: allowed,
    omissions: false,
    omittedCount: 0,
    omittedReferences: [],
    localCount: 0,
    notInspectedCount: 0,
  };
}

export function buildReferenceReport(allowed: boolean, a: ReferenceAnalysis): ReferenceReport {
  const base = emptyReferenceReport(allowed);
  base.localCount = a.local.length;
  base.notInspectedCount = a.notInspected.length;
  if (allowed) {
    base.recognisedCount = a.found.length;
    base.recognisedReferences = a.found.slice(0, REFERENCE_LIST_CAP);
  } else {
    base.omissions = a.found.length > 0;
    base.omittedCount = a.found.length;
    base.omittedReferences = a.found.slice(0, REFERENCE_LIST_CAP);
  }
  return base;
}

/** One stderr/CLI line for a report, or null when there is nothing to say. */
export function describeReferenceReport(r: ReferenceReport): string | null {
  if (r.networkReferencesAllowed) {
    if (r.scanFailed) return 'network references allowed; the scan for them failed, so the renderer may contact locations the document names';
    if ((r.recognisedCount ?? 0) > 0) {
      return `network references allowed; the renderer may contact ${r.recognisedCount} location${r.recognisedCount === 1 ? '' : 's'} found in the document (the scan is not exhaustive)`;
    }
    return null;
  }
  if (r.omittedCount > 0) {
    return `${r.omittedCount} network reference${r.omittedCount === 1 ? ' was' : 's were'} not loaded (render copy only; the file is unchanged). Use --allow-network-references to load them`;
  }
  return null;
}
