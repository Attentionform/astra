/**
 * Namespace prefix canonicalisation (DN3, Phase 0 prototype; not wired into any code path).
 *
 * A part may bind a namespace to any prefix. The engine reads parts by the conventional prefix, so a part is
 * rewritten here to use the conventional prefix for every namespace in `CANONICAL_PREFIXES`.
 *
 * Design rules that keep the meaning of the part:
 *  - Text-level rewrite: only element names, attribute names and namespace declarations in start and end tags
 *    change. Text, CDATA, comments, processing instructions, entities and whitespace keep their bytes.
 *  - Every original declaration stays where it was. The canonical declaration is added next to it. A value
 *    that names an old prefix (`mc:Ignorable="x"`, `xsi:type="x:T"`, anything an extension wrote) therefore still
 *    resolves to the same namespace without ASTra having to understand it.
 *  - The canonical prefix must resolve to the right namespace in scope at every renamed node, or the part is
 *    refused. Nothing is written by this module; the caller decides what to do with the result.
 *  - `verifyEquivalent` is an independent check (it parses both texts with the hardened parser and compares
 *    namespace-resolved forms and binding scopes). It does not trust the rewriter.
 */
import { DOMParser } from './dom-parser';
import { decodeXmlPart } from './part-text';
import { decodeXmlEntities } from './xml-entities';

export const CANONICAL_PREFIXES: Readonly<Record<string, string>> = {
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'urn:schemas-microsoft-com:vml': 'v',
  'urn:schemas-microsoft-com:office:office': 'o',
  'urn:schemas-microsoft-com:office:word': 'w10',
  'http://schemas.microsoft.com/office/word/2010/wordml': 'w14',
  'http://schemas.microsoft.com/office/word/2012/wordml': 'w15',
};

const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
/** Markup-compatibility attributes whose value is a whitespace-separated list of prefixes. */
const MC_PREFIX_LISTS = new Set(['Ignorable', 'MustUnderstand']);
/** Markup-compatibility attributes whose value is a whitespace-separated list of qualified names. */
const MC_QNAME_LISTS = new Set(['ProcessContent', 'PreserveElements', 'PreserveAttributes']);

export class PrefixRefusal extends Error {
  readonly code = 'E_VALIDATION';
  constructor(message: string, readonly reason: 'conflict' | 'unresolved' | 'malformed') {
    super(message);
  }
}

export interface Rename {
  from: string;
  to: string;
  uri: string;
}

export interface ScanResult {
  /** Table namespaces bound to a prefix other than the canonical one (or to the default namespace). */
  nonCanonical: Rename[];
}

const NAME_CHARS = '[\\p{L}_][\\p{L}\\p{N}._:\\u00B7-]*';
const DECLARATION = /\sxmlns(?::([\p{L}_][\p{L}\p{N}._\u00B7-]*))?\s*=\s*(?:"([^"]*)"|'([^']*)')/gu;
const STRIP_OPAQUE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g;
/** One start, end or empty tag, quote-aware (a `>` inside an attribute value does not end it). */
const SCAN_TAG = /<(\/?)([^\s<>/!?=]+)((?:[^<>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
const DEFAULT_DECL = /\sxmlns\s*=\s*(?:"([^"]*)"|'([^']*)')/;
const ATTR_PREFIX = /\s([^\s=:<>"']+):[^\s=:<>"']+\s*=/g;

/**
 * Cheap text scan, no parse: null when nothing needs a rewrite, otherwise the bindings that do. A table namespace
 * bound to a non-canonical prefix needs a rewrite only when some tag still writes a name with that prefix, or (for
 * the default namespace) an unprefixed element sits in that default namespace. A part that was rewritten before
 * keeps its original declaration (so values that name it still resolve); on the next open that declaration is
 * unused, so the part is left alone and the second open reports and rewrites nothing. Comments, CDATA and
 * processing instructions are ignored throughout.
 */
export function scanPart(source: string): ScanResult | null {
  const declarationsIn = (text: string): Map<string, Rename> => {
    const out = new Map<string, Rename>();
    DECLARATION.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = DECLARATION.exec(text)) !== null) {
      const prefix = m[1] ?? '';
      // A namespace written with a character reference (`&#104;ttp://...`) is the same namespace once parsed.
      const uri = decodeXmlEntities(m[2] ?? m[3] ?? '');
      const canonical = CANONICAL_PREFIXES[uri];
      if (canonical === undefined || prefix === canonical) continue;
      out.set(`${prefix}|${uri}`, { from: prefix, to: canonical, uri });
    }
    return out;
  };
  // Fast path for the usual case: no non-canonical declaration anywhere in the raw text, so nothing else to read.
  if (declarationsIn(source).size === 0) return null;
  const xml = source.replace(STRIP_OPAQUE, '');
  const found = declarationsIn(xml);
  if (found.size === 0) return null;

  const usedPrefixes = new Set<string>();
  const usedDefaults = new Set<string>();
  const defaults: string[] = [];
  let currentDefault = '';
  SCAN_TAG.lastIndex = 0;
  for (let t = SCAN_TAG.exec(xml); t !== null; t = SCAN_TAG.exec(xml)) {
    const closing = t[1] === '/';
    if (closing) {
      currentDefault = defaults.pop() ?? '';
      continue;
    }
    const attrs = t[3] ?? '';
    const selfClosing = t[4] === '/';
    defaults.push(currentDefault);
    const decl = DEFAULT_DECL.exec(attrs);
    if (decl) currentDefault = decodeXmlEntities(decl[1] ?? decl[2] ?? '');
    const name = t[2];
    const colon = name.indexOf(':');
    if (colon < 0) usedDefaults.add(currentDefault);
    else usedPrefixes.add(name.slice(0, colon));
    const bare = attrs.replace(/"[^"]*"|'[^']*'/g, '""');
    ATTR_PREFIX.lastIndex = 0;
    for (let a = ATTR_PREFIX.exec(bare); a !== null; a = ATTR_PREFIX.exec(bare)) usedPrefixes.add(a[1]);
    if (selfClosing) currentDefault = defaults.pop() ?? '';
  }
  const used = [...found.values()].filter((r) => (r.from === '' ? usedDefaults.has(r.uri) : usedPrefixes.has(r.from)));
  return used.length === 0 ? null : { nonCanonical: used };
}

const TAG = new RegExp(`<(/?)(${NAME_CHARS})((?:\\s+${NAME_CHARS}\\s*=\\s*(?:"[^"]*"|'[^']*'))*)\\s*(/?)>`, 'uy');
const ATTR = new RegExp(`(\\s+)(${NAME_CHARS})(\\s*=\\s*)(?:"([^"]*)"|'([^']*)')`, 'gu');

/** The text outside tags that must not be tokenised: comments, CDATA, processing instructions. */
const OPAQUE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/y;

function splitName(qname: string): { prefix: string; local: string } {
  const i = qname.indexOf(':');
  return i < 0 ? { prefix: '', local: qname } : { prefix: qname.slice(0, i), local: qname.slice(i + 1) };
}

/**
 * A scope is the declarations of one element plus a link to the enclosing scope. An element without declarations
 * shares its parent's scope object, so a part nested 100,000 deep does not copy or walk 100,000 maps. Lookups cache
 * their answer on every scope they pass, so each (scope, prefix) pair is resolved once.
 */
interface Scope {
  decl: Map<string, string>;
  parent: Scope | null;
  cache: Map<string, string | null>;
}

const newScope = (parent: Scope | null): Scope => ({ decl: new Map(), parent, cache: new Map() });

function lookup(stack: Scope[], prefix: string): string | undefined {
  let cur: Scope | null = stack.length > 0 ? stack[stack.length - 1] : null;
  const path: Scope[] = [];
  let result: string | undefined;
  while (cur !== null) {
    const cached = cur.cache.get(prefix);
    if (cached !== undefined) {
      result = cached === null ? undefined : cached;
      break;
    }
    const declared = cur.decl.get(prefix);
    if (declared !== undefined) {
      result = declared;
      cur.cache.set(prefix, declared);
      break;
    }
    path.push(cur);
    cur = cur.parent;
  }
  for (const scope of path) scope.cache.set(prefix, result ?? null);
  return result;
}

function declare(scope: Scope, prefix: string, uri: string): void {
  scope.decl.set(prefix, uri);
  scope.cache.set(prefix, uri);
}

/**
 * Rewrites `xml` so every table namespace uses its canonical prefix. Returns the new text and the renames
 * applied. Throws `PrefixRefusal` (nothing is written anywhere) on a conflict or an unresolvable value.
 */
export function canonicalisePrefixes(xml: string, partName = 'part', alreadyScanned = false): { xml: string; renames: Rename[] } {
  if (!alreadyScanned && scanPart(xml) === null) return { xml, renames: [] };
  const out: string[] = [];
  const renames = new Map<string, Rename>();
  const scopeStack: Scope[] = []; // declarations as they are in the INPUT, per open element
  const planned: Scope[] = []; // declarations as they will be in the OUTPUT, per open element
  const openNames: Array<{ input: string; output: string }> = [];
  let pos = 0;

  const refuse = (reason: PrefixRefusal['reason'], detail: string): never => {
    throw new PrefixRefusal(
      `E_VALIDATION: ${partName} cannot be canonicalised safely: ${detail}. Open the file in Word, save it again, then retry.`,
      reason
    );
  };

  while (pos < xml.length) {
    const lt = xml.indexOf('<', pos);
    if (lt < 0) {
      out.push(xml.slice(pos));
      break;
    }
    out.push(xml.slice(pos, lt));
    OPAQUE.lastIndex = lt;
    const opaque = OPAQUE.exec(xml);
    if (opaque) {
      out.push(opaque[0]);
      pos = lt + opaque[0].length;
      continue;
    }
    if (xml.startsWith('<!', lt)) refuse('malformed', 'it has a document type declaration');
    TAG.lastIndex = lt;
    const m = TAG.exec(xml);
    if (!m) return refuse('malformed', `unreadable markup near offset ${lt}`);
    pos = lt + m[0].length;
    const close = m[1] === '/';
    const qname = m[2];
    const attrText = m[3] ?? '';
    const selfClose = m[4] === '/';

    if (close) {
      const open = openNames.pop();
      if (!open || open.input !== qname) refuse('malformed', `end tag </${qname}> does not match`);
      out.push(`</${(open as { output: string }).output}>`);
      scopeStack.pop();
      planned.pop();
      continue;
    }

    // Collect the attributes and this element's declarations.
    interface A { ws: string; name: string; eq: string; quote: string; value: string }
    const attrs: A[] = [];
    ATTR.lastIndex = 0;
    let am: RegExpExecArray | null;
    while ((am = ATTR.exec(attrText)) !== null) {
      const dq = am[4] !== undefined;
      attrs.push({ ws: am[1], name: am[2], eq: am[3], quote: dq ? '"' : "'", value: dq ? am[4] : am[5] });
    }
    const declaredHere: Array<[string, string]> = [];
    for (const a of attrs) {
      if (a.name === 'xmlns') declaredHere.push(['', decodeXmlEntities(a.value)]);
      else if (a.name.startsWith('xmlns:')) declaredHere.push([a.name.slice(6), decodeXmlEntities(a.value)]);
    }
    const extraDecls: string[] = [];
    const parentIn = scopeStack.length > 0 ? scopeStack[scopeStack.length - 1] : null;
    const parentOut = planned.length > 0 ? planned[planned.length - 1] : null;
    // No declarations: share the enclosing scope objects (see Scope).
    const inScope = declaredHere.length > 0 ? newScope(parentIn) : (parentIn ?? newScope(null));
    const outScope = declaredHere.length > 0 ? newScope(parentOut) : (parentOut ?? newScope(null));
    for (const [p, u] of declaredHere) {
      declare(inScope, p, u);
      declare(outScope, p, u);
    }
    scopeStack.push(inScope);
    planned.push(outScope);
    // Add the canonical declaration next to each declaration of a table namespace under another prefix.
    for (const [p, u] of declaredHere) {
      const canonical = CANONICAL_PREFIXES[u];
      if (canonical === undefined || p === canonical) continue;
      const present = lookup(planned, canonical);
      if (present !== undefined && present !== u) {
        refuse('conflict', `the prefix "${canonical}" is bound to "${present}" where "${u}" (prefix "${p || '(default)'}") needs it`);
      }
      if (present === undefined) {
        declare(outScope, canonical, u);
        extraDecls.push(` xmlns:${canonical}="${u}"`);
      }
      const key = `${p}|${u}`;
      if (!renames.has(key)) renames.set(key, { from: p, to: canonical, uri: u });
    }

    /** The canonical qualified name for a prefix in INPUT scope; the original when no rewrite is needed. */
    const rewriteName = (qn: string, isAttribute: boolean): string => {
      const { prefix, local } = splitName(qn);
      if (isAttribute && prefix === '') return qn; // unprefixed attributes have no namespace
      if (prefix === 'xmlns' || prefix === 'xml' || qn === 'xmlns') return qn;
      const uri = lookup(scopeStack, prefix);
      if (uri === undefined) {
        if (prefix === '') return qn; // no default namespace in scope: no namespace
        return refuse('unresolved', `the prefix "${prefix}" is not declared (on <${qname}>)`);
      }
      const canonical = CANONICAL_PREFIXES[uri];
      if (canonical === undefined || canonical === prefix) return qn;
      const bound = lookup(planned, canonical);
      if (bound !== uri) refuse('conflict', `the prefix "${canonical}" would not resolve to "${uri}" on <${qname}>`);
      return `${canonical}:${local}`;
    };

    const elementParts = splitName(qname);
    const elementIsMc = lookup(scopeStack, elementParts.prefix) === MC_NS;
    const rewriteMcValue = (attrQName: string, value: string): string => {
      const { prefix, local } = splitName(attrQName);
      const toCanonical = (p: string): string => {
        const uri = lookup(scopeStack, p);
        // A list that names a prefix declared nowhere is left as written: it was readable before and stays so.
        if (uri === undefined) return p;
        const canonical = CANONICAL_PREFIXES[uri];
        if (canonical === undefined || canonical === p) return p;
        return lookup(planned, canonical) === uri ? canonical : p;
      };
      const prefixList = (v: string): string => v.split(/(\s+)/).map((tok) => (/^\s*$/.test(tok) ? tok : toCanonical(tok))).join('');
      // mc:Ignorable and mc:MustUnderstand carry prefix names; Requires (no namespace) on mc:Choice does too.
      if (prefix !== '' && lookup(scopeStack, prefix) === MC_NS) {
        if (MC_PREFIX_LISTS.has(local)) return prefixList(value);
        if (MC_QNAME_LISTS.has(local)) {
          return value
            .split(/(\s+)/)
            .map((tok) => {
              if (/^\s*$/.test(tok)) return tok;
              const q = splitName(tok);
              return q.prefix === '' ? tok : `${toCanonical(q.prefix)}:${q.local}`;
            })
            .join('');
        }
      }
      if (prefix === '' && local === 'Requires' && elementIsMc && elementParts.local === 'Choice') return prefixList(value);
      return value;
    };

    const newName = rewriteName(qname, false);
    const seenAttributes = new Set<string>();
    const rebuilt = attrs
      .map((a) => {
        if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) return `${a.ws}${a.name}${a.eq}${a.quote}${a.value}${a.quote}`;
        const n = rewriteName(a.name, true);
        if (seenAttributes.has(n)) refuse('conflict', `two attributes of <${qname}> would both be written as "${n}"`);
        seenAttributes.add(n);
        const v = rewriteMcValue(a.name, a.value);
        return `${a.ws}${n}${a.eq}${a.quote}${v}${a.quote}`;
      })
      .join('');
    out.push(`<${newName}${rebuilt}${extraDecls.join('')}${selfClose ? '/' : ''}>`);
    if (selfClose) {
      scopeStack.pop();
      planned.pop();
    } else {
      openNames.push({ input: qname, output: newName });
    }
  }
  if (openNames.length > 0) refuse('malformed', `<${openNames[openNames.length - 1].input}> is not closed`);
  return { xml: out.join(''), renames: [...renames.values()] };
}

// ---------------------------------------------------------------------------------------------------------------
// Independent verification
// ---------------------------------------------------------------------------------------------------------------

interface Node1 {
  nodeType: number;
  namespaceURI?: string | null;
  localName?: string | null;
  attributes?: ArrayLike<{ name: string; value: string; namespaceURI?: string | null; localName?: string | null }>;
  childNodes?: ArrayLike<Node1>;
  data?: string;
  nodeValue?: string | null;
  tagName?: string;
}

function parseStrict(xml: string): Node1 {
  const fail = (message: string): never => {
    throw new PrefixRefusal(`unreadable XML while verifying: ${message}`, 'malformed');
  };
  const doc = new DOMParser({ errorHandler: { warning: () => undefined, error: fail, fatalError: fail } }).parseFromString(xml, 'text/xml') as unknown as {
    documentElement: Node1;
  };
  if (!doc.documentElement) fail('no root element');
  return doc.documentElement;
}

interface Walked {
  /** Namespace-resolved description of every node, in document order. */
  resolved: string[];
  /** For every element, the prefix to namespace bindings in scope. */
  scopes: Array<Map<string, string>>;
}

/** A real part binds a few dozen prefixes. A scope far beyond that is hostile, and comparing it per element would not terminate in useful time. */
const MAX_BINDINGS_IN_SCOPE = 4096;

function walk(root: Node1): Walked {
  const resolved: string[] = [];
  const scopes: Array<Map<string, string>> = [];
  // Iterative on purpose: a document can nest tens of thousands of elements and recursion would overflow the stack.
  type Frame = { node: Node1; scope: Map<string, string>; close?: string };
  const stack: Frame[] = [{ node: root, scope: new Map() }];
  while (stack.length > 0) {
    const frame = stack.pop() as Frame;
    if (frame.close !== undefined) {
      resolved.push(frame.close);
      continue;
    }
    const node = frame.node;
    if (node.nodeType === 3 || node.nodeType === 4) {
      resolved.push(`#${node.nodeType}:${node.data ?? node.nodeValue ?? ''}`);
      continue;
    }
    if (node.nodeType !== 1) {
      resolved.push(`#${node.nodeType}:${node.nodeValue ?? node.data ?? ''}`);
      continue;
    }
    const attrs: Array<{ name: string; value: string; namespaceURI?: string | null; localName?: string | null }> = [];
    for (let i = 0; i < (node.attributes?.length ?? 0); i++) attrs.push((node.attributes as ArrayLike<(typeof attrs)[number]>)[i]);
    // Copy the scope only when this element declares something; otherwise share the parent's (see `Scope`).
    const declares = attrs.some((a) => a.name === 'xmlns' || a.name.startsWith('xmlns:'));
    const scope = declares ? new Map(frame.scope) : frame.scope;
    if (declares) {
      for (const a of attrs) {
        if (a.name === 'xmlns') scope.set('', a.value);
        else if (a.name.startsWith('xmlns:')) scope.set(a.name.slice(6), a.value);
      }
      if (scope.size > MAX_BINDINGS_IN_SCOPE) {
        throw new PrefixRefusal(`unreadable XML while verifying: more than ${MAX_BINDINGS_IN_SCOPE} namespace bindings in scope`, 'malformed');
      }
    }
    scopes.push(scope);
    const uriOf = (prefix: string): string => scope.get(prefix) ?? '';
    const own = splitName(node.tagName ?? '');
    const elementKey = `{${uriOf(own.prefix)}}${own.local}`;
    const isMc = uriOf(own.prefix) === MC_NS;
    const attrKeys: string[] = [];
    for (const a of attrs) {
      if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) continue;
      const n = splitName(a.name);
      const key = a.name.startsWith('xml:') ? a.name : n.prefix === '' ? n.local : `{${uriOf(n.prefix)}}${n.local}`;
      let value = a.value;
      // A value that lists prefixes or qualified names is compared by the namespaces it names, not by its text.
      if (n.prefix !== '' && uriOf(n.prefix) === MC_NS && MC_PREFIX_LISTS.has(n.local)) {
        value = a.value.split(/\s+/).filter(Boolean).map((p) => `{${uriOf(p)}}`).join(' ');
      } else if (n.prefix !== '' && uriOf(n.prefix) === MC_NS && MC_QNAME_LISTS.has(n.local)) {
        value = a.value.split(/\s+/).filter(Boolean).map((q) => {
          const sq = splitName(q);
          return `{${uriOf(sq.prefix)}}${sq.local}`;
        }).join(' ');
      } else if (isMc && own.local === 'Choice' && n.prefix === '' && n.local === 'Requires') {
        value = a.value.split(/\s+/).filter(Boolean).map((p) => `{${uriOf(p)}}`).join(' ');
      }
      attrKeys.push(`${key}=${value}`);
    }
    attrKeys.sort();
    resolved.push(`<${elementKey} ${attrKeys.join(' ')}>`);
    // Close marker first onto the stack, then the children in reverse so they pop in document order.
    stack.push({ node, scope, close: `</${elementKey}>` });
    const kids = node.childNodes;
    for (let i = (kids?.length ?? 0) - 1; i >= 0; i--) stack.push({ node: (kids as ArrayLike<Node1>)[i], scope });
  }
  return { resolved, scopes };
}

export interface Equivalence {
  ok: boolean;
  /** First difference, for the refusal text and for tests. */
  problem?: string;
}

/**
 * Independent check that `after` means what `before` means: same namespace-resolved node sequence (names,
 * attribute values, text; prefix-list values compared by the namespaces they name), and every binding that was
 * in scope at an element before is still in scope there with the same namespace. Not a proof of all XML meaning:
 * a value that mentions a prefix in a syntax this module does not know keeps resolving only because the old
 * binding is kept, and that is what the binding check asserts.
 */
export function verifyEquivalent(before: string, after: string): Equivalence {
  const a = walk(parseStrict(before));
  const b = walk(parseStrict(after));
  if (a.resolved.length !== b.resolved.length) return { ok: false, problem: `node count ${a.resolved.length} became ${b.resolved.length}` };
  for (let i = 0; i < a.resolved.length; i++) {
    if (a.resolved[i] !== b.resolved[i]) return { ok: false, problem: `node ${i}: ${a.resolved[i].slice(0, 120)} became ${b.resolved[i].slice(0, 120)}` };
  }
  if (a.scopes.length !== b.scopes.length) return { ok: false, problem: 'element count differs' };
  for (let i = 0; i < a.scopes.length; i++) {
    // An element that declares nothing shares its parent's scope object on both sides: already compared.
    if (i > 0 && a.scopes[i] === a.scopes[i - 1] && b.scopes[i] === b.scopes[i - 1]) continue;
    for (const [prefix, uri] of a.scopes[i]) {
      if (b.scopes[i].get(prefix) !== uri) {
        return { ok: false, problem: `element ${i}: the binding ${prefix || '(default)'} -> ${uri} is no longer in scope` };
      }
    }
  }
  return { ok: true };
}

/** Canonicalises and proves the result. Throws `PrefixRefusal` if the check fails. */
export function canonicalisePart(xml: string, partName = 'part', alreadyScanned = false): { xml: string; renames: Rename[] } {
  const result = canonicalisePrefixes(xml, partName, alreadyScanned);
  if (result.renames.length === 0) return result;
  const check = verifyEquivalent(xml, result.xml);
  if (!check.ok) {
    throw new PrefixRefusal(`E_VALIDATION: ${partName} cannot be canonicalised safely: the rewrite changed its meaning (${check.problem}). Open the file in Word, save it again, then retry.`, 'conflict');
  }
  return result;
}

export interface IngestResult {
  /** Replacement text for each part that needed a rewrite. Parts that were already canonical are absent. */
  changed: Map<string, string>;
  /** Parts that were rewritten, with the renames applied. For the open summary. */
  report: Array<{ part: string; renames: Rename[] }>;
}

/**
 * Staging ingress: canonicalises every part in `parts` (name to decoded text). Nothing is written by this
 * function. It either returns the replacements for all parts that need one, or throws `PrefixRefusal` for the
 * first part that cannot be rewritten safely; the caller writes the replacements only after it returns.
 * `utf16` names parts stored as UTF-16: a part like that is refused instead of rewritten, because the rewrite
 * would write it back as UTF-8 under a declaration that says otherwise.
 */
export function ingestPrefixes(parts: Iterable<[string, string]>, utf16: ReadonlySet<string> = new Set()): IngestResult {
  const changed = new Map<string, string>();
  const report: IngestResult['report'] = [];
  for (const [name, text] of parts) {
    if (scanPart(text) === null) continue;
    if (utf16.has(name)) {
      throw new PrefixRefusal(
        `E_VALIDATION: ${name} is stored as UTF-16 and uses a prefix other than the usual one; ASTra does not rewrite it. Open the file in Word, save it again, then retry.`,
        'malformed'
      );
    }
    const result = canonicalisePart(text, name, true);
    // The scan is a cheap text test; when the parser-level rewrite finds nothing to rename, leave the part alone.
    if (result.renames.length === 0) continue;
    changed.set(name, result.xml);
    report.push({ part: name, renames: result.renames });
  }
  return { changed, report };
}

export const looksUtf16 = (data: Buffer): boolean =>
  data.length >= 2 && ((data[0] === 0xff && data[1] === 0xfe) || (data[0] === 0xfe && data[1] === 0xff) || data[0] === 0 || data[1] === 0);

/**
 * `ingestPrefixes` over a map of part bytes (a package snapshot, the entries a sanitizer is about to read). Returns
 * the same map when nothing needs a rewrite, else a copy with the rewritten parts. Throws `PrefixRefusal`.
 * `skip` names entries that are not read as parts (rogue or blocked names).
 */
export function canonicaliseBufferMap(
  files: Map<string, Buffer>,
  skip: (name: string) => boolean = () => false
): { files: Map<string, Buffer>; report: IngestResult['report'] } {
  const parts: Array<[string, string]> = [];
  const utf16 = new Set<string>();
  for (const [name, data] of files) {
    if (!name.toLowerCase().endsWith('.xml') || skip(name)) continue;
    if (looksUtf16(data)) utf16.add(name);
    parts.push([name, decodeXmlPart(data)]);
  }
  const result = ingestPrefixes(parts, utf16);
  if (result.changed.size === 0) return { files, report: [] };
  const copy = new Map(files);
  for (const [name, text] of result.changed) copy.set(name, Buffer.from(text, 'utf8'));
  return { files: copy, report: result.report };
}
