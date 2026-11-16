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

const DECLARATION = /\sxmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/**
 * Cheap text scan, no parse: null when every table namespace is bound to its canonical prefix everywhere,
 * otherwise the bindings that need a rewrite. Also reports a canonical prefix bound to a different namespace
 * only through the rewrite (a conflict is found there, where scope is known).
 */
export function scanPart(xml: string): ScanResult | null {
  const found = new Map<string, Rename>();
  DECLARATION.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = DECLARATION.exec(xml)) !== null) {
    const prefix = m[1] ?? '';
    const uri = m[2] ?? m[3] ?? '';
    const canonical = CANONICAL_PREFIXES[uri];
    if (canonical === undefined || prefix === canonical) continue;
    found.set(`${prefix}|${uri}`, { from: prefix, to: canonical, uri });
  }
  return found.size === 0 ? null : { nonCanonical: [...found.values()] };
}

const TAG = /<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/y;
const ATTR = /(\s+)([A-Za-z_][\w.:-]*)(\s*=\s*)(?:"([^"]*)"|'([^']*)')/g;

/** The text outside tags that must not be tokenised: comments, CDATA, processing instructions. */
const OPAQUE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/y;

function splitName(qname: string): { prefix: string; local: string } {
  const i = qname.indexOf(':');
  return i < 0 ? { prefix: '', local: qname } : { prefix: qname.slice(0, i), local: qname.slice(i + 1) };
}

type Scope = Map<string, string>;

function lookup(scope: Scope[], prefix: string): string | undefined {
  for (let i = scope.length - 1; i >= 0; i--) {
    const v = scope[i].get(prefix);
    if (v !== undefined) return v;
  }
  return undefined;
}

/**
 * Rewrites `xml` so every table namespace uses its canonical prefix. Returns the new text and the renames
 * applied. Throws `PrefixRefusal` (nothing is written anywhere) on a conflict or an unresolvable value.
 */
export function canonicalisePrefixes(xml: string, partName = 'part'): { xml: string; renames: Rename[] } {
  if (scanPart(xml) === null) return { xml, renames: [] };
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
    const inScope: Scope = new Map();
    const outScope: Scope = new Map();
    const extraDecls: string[] = [];
    for (const a of attrs) {
      if (a.name === 'xmlns') inScope.set('', a.value);
      else if (a.name.startsWith('xmlns:')) inScope.set(a.name.slice(6), a.value);
    }
    scopeStack.push(inScope);
    for (const [p, u] of inScope) outScope.set(p, u);
    planned.push(outScope);
    // Add the canonical declaration next to each declaration of a table namespace under another prefix.
    for (const [p, u] of inScope) {
      const canonical = CANONICAL_PREFIXES[u];
      if (canonical === undefined || p === canonical) continue;
      const present = lookup(planned, canonical);
      if (present !== undefined && present !== u) {
        refuse('conflict', `the prefix "${canonical}" is bound to "${present}" where "${u}" (prefix "${p || '(default)'}") needs it`);
      }
      if (present === undefined) {
        outScope.set(canonical, u);
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
        if (uri === undefined) return refuse('unresolved', `${attrQName} names the undeclared prefix "${p}"`);
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
    const rebuilt = attrs
      .map((a) => {
        if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) return `${a.ws}${a.name}${a.eq}${a.quote}${a.value}${a.quote}`;
        const n = rewriteName(a.name, true);
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

function walk(root: Node1): Walked {
  const resolved: string[] = [];
  const scopes: Array<Map<string, string>> = [];
  const visit = (node: Node1, parent: Map<string, string>): void => {
    if (node.nodeType === 3 || node.nodeType === 4) {
      resolved.push(`#${node.nodeType}:${node.data ?? node.nodeValue ?? ''}`);
      return;
    }
    if (node.nodeType !== 1) {
      resolved.push(`#${node.nodeType}:${node.nodeValue ?? node.data ?? ''}`);
      return;
    }
    const scope = new Map(parent);
    const attrs: Array<{ name: string; value: string; namespaceURI?: string | null; localName?: string | null }> = [];
    for (let i = 0; i < (node.attributes?.length ?? 0); i++) attrs.push((node.attributes as ArrayLike<(typeof attrs)[number]>)[i]);
    for (const a of attrs) {
      if (a.name === 'xmlns') scope.set('', a.value);
      else if (a.name.startsWith('xmlns:')) scope.set(a.name.slice(6), a.value);
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
          const s = splitName(q);
          return `{${uriOf(s.prefix)}}${s.local}`;
        }).join(' ');
      } else if (isMc && own.local === 'Choice' && n.prefix === '' && n.local === 'Requires') {
        value = a.value.split(/\s+/).filter(Boolean).map((p) => `{${uriOf(p)}}`).join(' ');
      }
      attrKeys.push(`${key}=${value}`);
    }
    attrKeys.sort();
    resolved.push(`<${elementKey} ${attrKeys.join(' ')}>`);
    for (let i = 0; i < (node.childNodes?.length ?? 0); i++) visit((node.childNodes as ArrayLike<Node1>)[i], scope);
    resolved.push(`</${elementKey}>`);
  };
  visit(root, new Map());
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
    for (const [prefix, uri] of a.scopes[i]) {
      if (b.scopes[i].get(prefix) !== uri) {
        return { ok: false, problem: `element ${i}: the binding ${prefix || '(default)'} -> ${uri} is no longer in scope` };
      }
    }
  }
  return { ok: true };
}

/** Canonicalises and proves the result. Throws `PrefixRefusal` if the check fails. */
export function canonicalisePart(xml: string, partName = 'part'): { xml: string; renames: Rename[] } {
  const result = canonicalisePrefixes(xml, partName);
  if (result.renames.length === 0) return result;
  const check = verifyEquivalent(xml, result.xml);
  if (!check.ok) {
    throw new PrefixRefusal(`E_VALIDATION: ${partName} cannot be canonicalised safely: the rewrite changed its meaning (${check.problem}). Open the file in Word, save it again, then retry.`, 'conflict');
  }
  return result;
}
