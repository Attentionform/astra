/**
 * Central case-insensitive XXE / DTD detector for every XML ingress point.
 *
 * Single pattern, single predicate, single throwing guard. Each ingress calls
 * `assertNoXxe` (or checks `containsForbiddenMarkup`) BEFORE handing
 * attacker-influenced bytes to a parser or a zip writer. Fail-closed: a hit
 * throws E_VALIDATION with a remedy; nothing is dropped silently and nothing
 * is auto-repaired.
 *
 * Detection boundary (verified by execution, not assumed):
 * - Matches literal `<` + one or more `!` + optional whitespace +
 *   `DOCTYPE` / `ENTITY`, case-insensitively. The `!+` run is load-bearing:
 *   lenient parsers (verified against @xmldom/xmldom) normalize adjacent
 *   multi-bang opens (`<!!DOCTYPE`, `<!!!DOCTYPE`, `<!!!!DOCTYPE`) into a
 *   real doctype node with no parsererror, so a single-`!` pattern lets them
 *   through. Whitespace-mixed multi-bang (`<!! DOCTYPE`, `<! !DOCTYPE`)
 *   escapes to inert text at the XML layer yet is still blocked fail-closed —
 *   no well-formed OOXML part ever carries a `!!` run.
 * - The whitespace arm blocks tab/newline-split evasions such as
 *   `<!\nDOCTYPE`, which lenient XML parsers still accept as markup.
 * - Scans the RAW input only. Encoded lookalikes (`&lt;!DOCTYPE`,
 *   `&#60;!ENTITY`, `%3C!DOCTYPE`, ...) decode to element TEXT at the XML
 *   layer, never to markup, so they are inert and intentionally NOT flagged.
 *   Decoding before scanning would false-positive on legit math and prose
 *   (for example `a &lt; b`) — that over-blocking is the documented line.
 * - Well-formed OOXML never carries a literal `<!DOCTYPE` / `<!ENTITY`
 *   (a text `<` is always escaped as `&lt;`), so fail-closed blocking breaks
 *   no legit document. `<!--` comments and `<![CDATA[` sections carry no
 *   DOCTYPE/ENTITY keyword and pass untouched.
 * - Comment-split evasions (`<!<!--...-->DOCTYPE`, `<!DO<!--...-->CTYPE`,
 *   `<! --DOCTYPE`): verified by execution against @xmldom/xmldom — the
 *   split shape never forms a declaration node. It either escapes to text
 *   (`&lt;!...&gt;`) or the comment is dropped leaving a bare root with no
 *   DOCTYPE/ENTITY node and no entity expansion. Deliberately NOT matched:
 *   the raw `<!` + keyword pattern stays single so legit comment-heavy
 *   documents (TOC estimate notes, multi-comment prose) never false-positive.
 *   Covered by a passthrough test asserting literal serialization.
 * - Bare `<!ELEMENT`, `<!ATTLIST`, `<!NOTATION` without a DOCTYPE wrapper:
 *   verified by execution — xmldom escapes each to text (`&lt;!...&gt;`),
 *   emits no parsererror, creates no declaration node, and expands nothing.
 *   No gate change: a passthrough test pins the escape-to-text behavior.
 * - Conservative edge: a declaration keyword wrapped in CDATA or an XML
 *   comment still matches, because the string layer cannot see quoting.
 *   Word never emits such shapes; rejecting them fail-closed is intended.
 * - Storage decoding boundary: the engine reads every zip part with
 *   `toString('utf8')`. A UTF-16-encoded part therefore arrives as mojibake
 *   (every ASCII byte followed by NUL), which this pattern deliberately does
 *   NOT match — there is no UTF-16 detector in this slice. That evasion is
 *   still not silent: the mojibake never parses (the strict xmldom
 *   errorHandler throws in the constructor), so a UTF-16 part chokes loudly
 *   instead of loading. Pinned by test, not assumed.
 */

// One pattern for the whole codebase: literal markup open, one or more `!`
// (covers `<!!DOCTYPE`-style runs that lenient parsers still accept),
// optional inter-tag whitespace (covers `<!\nDOCTYPE` style splits), then
// the declaration keyword, case-insensitively.
const FORBIDDEN_MARKUP_RE = /<!+[\s\0]*(?:doctype|entity)/i;

/** Pure predicate: true when the raw input carries a DTD declaration attempt. */
export function containsForbiddenMarkup(xml: unknown): boolean {
  if (typeof xml !== 'string' || xml.length === 0) return false;
  return FORBIDDEN_MARKUP_RE.test(xml);
}

function contextLabel(context: string | undefined): string {
  return typeof context === 'string' && context.length > 0 ? context : 'XML part';
}

/** Builds the fail-closed E_VALIDATION message (kept uniform for all callers). */
export function buildXxeValidationMessage(context?: string): string {
  return (
    `E_VALIDATION: Security Error: XXE / DTD entity declarations are strictly prohibited` +
    ` in ${contextLabel(context)}.` +
    ` Remedy: Remove <!DOCTYPE and <!ENTITY declarations.`
  );
}

/**
 * Fail-closed guard: throws E_VALIDATION with a remedy when the raw input
 * carries a DTD declaration attempt. Never repairs, never swallows.
 */
export function assertNoXxe(xml: unknown, context = 'XML part'): void {
  if (containsForbiddenMarkup(xml)) {
    throw new Error(buildXxeValidationMessage(context));
  }
}

/** Compatibility alias for the previous sanitizer-local detector name. */
export const assertNoForbiddenMarkup = assertNoXxe;

// ── H3 XXE-adjacent vectors (D5) ─────────────────────────────────────────────
// The central DOCTYPE/ENTITY pattern above stays untouched (its boundary is
// pinned by tests). These SEPARATE predicates cover the adjacent vectors the
// audit probe confirmed sailing through ingest: XInclude, XSLT, the
// xml-stylesheet processing instruction, and SVG script-capable constructs
// (event-handler attributes, foreignObject). They are intentionally NOT
// folded into FORBIDDEN_MARKUP_RE: callers pick the fail-closed assert for
// their contract (refuse on compose ingest) or the strip helper for theirs
// (strip-loud on sanitize paths).

/** `<xi:include ...>` element open (any case, any attribute shape). */
const XINCLUDE_RE = /<\s*xi:include(?=[\s>/])/i;
/** XSLT namespace use: `<xsl:*>` element opens. */
const XSLT_ELEMENT_RE = /<\s*xsl:[\w-]+(?=[\s>/])/i;
/** XSLT namespace declaration (corroborating signal, same verdict). */
const XSLT_NS_RE = /http:\/\/www\.w3\.org\/1999\/XSL\/Transform/i;
/** `<?xml-stylesheet ...?>` processing instruction (not the `<?xml ...?>` declaration). */
const STYLESHEET_PI_RE = /<\?xml-stylesheet(?=[\s?])/i;

/** True when the raw input carries an XInclude directive. */
export function containsXInclude(xml: unknown): boolean {
  if (typeof xml !== 'string' || xml.length === 0) return false;
  return XINCLUDE_RE.test(xml);
}

/** True when the raw input carries XSLT (element use or namespace declaration). */
export function containsXslt(xml: unknown): boolean {
  if (typeof xml !== 'string' || xml.length === 0) return false;
  return XSLT_ELEMENT_RE.test(xml) || XSLT_NS_RE.test(xml);
}

/** True when the raw input carries an xml-stylesheet processing instruction. */
export function containsStylesheetPI(xml: unknown): boolean {
  if (typeof xml !== 'string' || xml.length === 0) return false;
  return STYLESHEET_PI_RE.test(xml);
}

/** True when any XXE-adjacent vector (XInclude, XSLT, stylesheet PI) is present. */
export function containsXxeAdjacent(xml: unknown): boolean {
  return containsXInclude(xml) || containsXslt(xml) || containsStylesheetPI(xml);
}

/**
 * Fail-closed guard for XXE-adjacent vectors: throws E_VALIDATION with a
 * remedy naming the construct. Compose ingest calls this (D5 refuse path).
 */
export function assertNoXxeAdjacent(xml: unknown, context = 'XML part'): void {
  if (containsXxeAdjacent(xml)) {
    const hits: string[] = [];
    if (containsXInclude(xml)) hits.push('xi:include');
    if (containsXslt(xml)) hits.push('xsl:');
    if (containsStylesheetPI(xml)) hits.push('xml-stylesheet PI');
    throw new Error(
      `E_VALIDATION: Security Error: XXE-adjacent constructs (${hits.join(', ')}) are strictly prohibited` +
        ` in ${contextLabel(context)}.` +
        ` Remedy: Remove XInclude directives, XSLT stylesheets, and xml-stylesheet processing instructions.`
    );
  }
}

// ── SVG script-capable constructs (D5) ───────────────────────────────────────

/** `<script ...>` open inside SVG (any case). */
const SVG_SCRIPT_RE = /<\s*script(?=[\s>/])/i;
/** `<foreignObject ...>` open inside SVG (any case). */
const SVG_FOREIGN_OBJECT_RE = /<\s*foreignObject(?=[\s>/])/i;
/** `<handler ...>` / `<listener ...>` (SMIL/SVG declarative event wiring). */
const SVG_HANDLER_ELEMENT_RE = /<\s*(handler|listener)(?=[\s>/])/i;

/**
 * Event-handler attributes (`onload=`, `onclick=`, ...) on any element open.
 * Quoted attribute VALUES are blanked before scanning so prose like
 * `<a title="turn on=off">` never false-positives — only real attribute
 * positions match.
 */
export function containsSvgEventHandler(svg: unknown): boolean {
  if (typeof svg !== 'string' || svg.length === 0) return false;
  const unquoted = (svg as string).replace(/"[^"]*"|'[^']*'/g, '""');
  return /<\s*[a-zA-Z][^>]*\bon[a-zA-Z]+\s*=/i.test(unquoted);
}

// ── F4 (R1 refined): refuse-only hrefs vs strip-only style ─────────────────
// Ingest-refuse = active content only: script, foreignObject,
// handler/listener, event-handler attrs (below) plus xi:include / xsl: /
// stylesheet-PI (via the adjacent asserts) and javascript:/vbscript: hrefs
// plus data:text/html hrefs. STRIP-ONLY (sanitize strips + reports loudly,
// ingest allows): `<style>` (Illustrator/Inkscape exports) and benign data:
// hrefs (data:image/* raster embeds, data:font/*). Both scans stay TAG-SCOPED
// (inside `<…>` only) so element TEXT prose
// (`<text>see javascript:foo</text>`) never trips them.
// Best-effort regex layer: quoted `>` inside an attribute value ends the tag
// scan early (documented miss, same class as the event-handler scan).
const SVG_URL_ATTR_RE = /(?:^|\s)(?:href|xlink:href|src)\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/i;
/** javascript:/vbscript: href values (refuse-only, any case, whitespace-tolerant). */
const SVG_REFUSE_JS_VBS_RE = /^\s*(?:javascript|vbscript)\s*:/i;
/**
 * data:text/html href values (refuse-only, any case, whitespace-tolerant,
 * optional `;charset` / `;base64` params). Benign data: hrefs
 * (data:image/* raster embeds, data:font/*) intentionally do NOT match —
 * they pass ingest and survive sanitize intact.
 */
const SVG_REFUSE_DATA_HTML_RE = /^\s*data\s*:\s*text\/html(?:[;\s,]|$)/i;
/** `<style ...>` open inside SVG (any case). The `[\s>/]` lookahead keeps
 *  `<styleSheet>` (spreadsheetML, different layer) out of the match.
 *  STRIP-ONLY: sanitize strips + reports, ingest allows (Illustrator/
 *  Inkscape exports carry presentation `<style>`). */
const SVG_STYLE_ELEMENT_RE = /<\s*style(?=[\s>/])/i;

/** True when a URL attribute (href/xlink:href/src) carries javascript:/vbscript:/data:text/html (refuse-only). data:image/*, data:font/*, https pass. */
export function containsSvgDangerousHref(svg: unknown): boolean {
  if (typeof svg !== 'string' || svg.length === 0) return false;
  const tags = (svg as string).match(/<[^<>]*>/g);
  if (!tags) return false;
  for (const tag of tags) {
    const m = SVG_URL_ATTR_RE.exec(tag);
    if (!m) continue;
    let v = m[1] ?? '';
    const q = v[0];
    if ((q === '"' || q === "'") && v.length >= 2 && v.endsWith(q)) v = v.slice(1, -1);
    if (SVG_REFUSE_JS_VBS_RE.test(v) || SVG_REFUSE_DATA_HTML_RE.test(v)) return true;
  }
  return false;
}

/** True when the SVG carries a `<style>` element (STRIP-ONLY: sanitize strips + reports, ingest allows). */
export function containsSvgStyleElement(svg: unknown): boolean {
  if (typeof svg !== 'string' || svg.length === 0) return false;
  return SVG_STYLE_ELEMENT_RE.test(svg);
}

/** True when the SVG carries refuse-only active content: script, foreignObject, handler elements, event attributes, or javascript:/vbscript:/data:text/html hrefs. `<style>` + data:image/font pass (strip-only / allowed). */
export function containsSvgActiveContent(svg: unknown): boolean {
  if (typeof svg !== 'string' || svg.length === 0) return false;
  return (
    SVG_SCRIPT_RE.test(svg) ||
    SVG_FOREIGN_OBJECT_RE.test(svg) ||
    SVG_HANDLER_ELEMENT_RE.test(svg) ||
    containsSvgEventHandler(svg) ||
    containsSvgDangerousHref(svg)
  );
}

/**
 * Fail-closed guard for SVG active content (D5 refuse path for compose
 * ingest): throws E_VALIDATION naming the construct class. Refuse-only:
 * script, foreignObject, event-handler attributes, javascript:/vbscript:/
 * data:text/html hrefs (plus xi:include/xsl:/stylesheet-PI via the adjacent
 * assert). `<style>` + data:image/font pass here; sanitize strips style loudly.
 */
export function assertNoSvgActiveContent(svg: unknown, context = 'SVG image'): void {
  if (containsSvgActiveContent(svg)) {
    throw new Error(
      `E_VALIDATION: Security Error: SVG active content (script, foreignObject, event-handler attributes, dangerous hrefs)` +
        ` is strictly prohibited in ${contextLabel(context)}.` +
        ` Remedy: Use a plain SVG without script, foreignObject, on* event attributes, or javascript:/vbscript:/data:text/html hrefs.`
    );
  }
}

/** One loud strip-report entry class for the sanitize-path helper below. */
export type SvgStripKind =
  | 'script'
  | 'foreignObject'
  | 'handler-element'
  | 'event-handler-attrs'
  | 'dangerous-href'
  | 'style-element'
  | 'xi:include'
  | 'xsl:stylesheet'
  | 'xml-stylesheet-PI';

export interface SvgStripResult {
  /** SVG with every script-capable construct removed (best-effort, regex layer). */
  cleaned: string;
  /** Loud report: one entry per removed construct class (counts, never content). */
  removed: SvgStripKind[];
}

/**
 * Sanitize-path contract (D5 strip path): removes script-capable constructs
 * from an SVG and reports WHAT was removed (kinds only, never the payload).
 * Best-effort regex layer — callers wire it into the XML-aware sanitizer
 * slice; a non-empty `removed` list is the loud disclosure. Never throws on
 * malformed input (returns the input unchanged with an empty report when
 * nothing matches).
 */
export function stripSvgActiveContent(svg: string): SvgStripResult {
  if (typeof svg !== 'string' || svg.length === 0) return { cleaned: svg, removed: [] };
  const removed: SvgStripKind[] = [];
  let out = svg;
  const drop = (re: RegExp, kind: SvgStripKind): void => {
    if (re.test(out)) {
      out = out.replace(re, '');
      if (!removed.includes(kind)) removed.push(kind);
    }
  };
  // Stylesheet PIs first (they sit outside the root element).
  drop(/<\?xml-stylesheet[\s?][^?]*\?>\s*/gi, 'xml-stylesheet-PI');
  // Paired elements (non-greedy to the first close; best-effort documented).
  drop(/<\s*script\b[^>]*>[\s\S]*?<\s*\/\s*script\s*>/gi, 'script');
  drop(/<\s*foreignObject\b[^>]*>[\s\S]*?<\s*\/\s*foreignObject\s*>/gi, 'foreignObject');
  drop(/<\s*(handler|listener)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, 'handler-element');
  // Self-closing leftovers of the same elements.
  drop(/<\s*script\b[^>]*\/\s*>/gi, 'script');
  drop(/<\s*foreignObject\b[^>]*\/\s*>/gi, 'foreignObject');
  drop(/<\s*xi:include\b[^>]*\/\s*>/gi, 'xi:include');
  drop(/<\s*xi:include\b[^>]*>[\s\S]*?<\s*\/\s*xi:include\s*>/gi, 'xi:include');
  // Any XSLT-namespace element (stylesheet pairs, value-of, template, ...):
  // the namespace itself is the payload, not just xsl:stylesheet.
  drop(/<\s*xsl:[\w-]+\b[^>]*>[\s\S]*?<\s*\/\s*xsl:[\w-]+\s*>/gi, 'xsl:stylesheet');
  drop(/<\s*xsl:[\w-]+\b[^>]*\/\s*>/gi, 'xsl:stylesheet');
  // Event-handler attributes on any element open (quoted values intact).
  if (containsSvgEventHandler(out)) {
    out = out.replace(/\s+on[a-zA-Z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '');
    if (!removed.includes('event-handler-attrs')) removed.push('event-handler-attrs');
  }
  // F4 (R1): <style> elements (paired + self-closing) — STRIP-ONLY.
  // Ingest allows (Illustrator/Inkscape exports); sanitize strips loudly.
  drop(/<\s*style\b[^>]*>[\s\S]*?<\s*\/\s*style\s*>/gi, 'style-element');
  drop(/<\s*style\b[^>]*\/\s*>/gi, 'style-element');
  // F4 (R1): refuse-only hrefs (javascript:/vbscript:/data:text/html) — the
  // attribute is dropped, the element survives. data:image/*, data:font/*,
  // https pass intact (raster embeds survive). Tag-scoped so text prose
  // never matches.
  {
    let strippedHref = false;
    out = out.replace(/<[^<>]*>/g, (tag) => {
      const m = SVG_URL_ATTR_RE.exec(tag);
      if (!m || m[1] === undefined) return tag;
      let v: string = m[1];
      const q = v[0];
      if ((q === '"' || q === "'") && v.length >= 2 && v.endsWith(q)) v = v.slice(1, -1);
      if (!(SVG_REFUSE_JS_VBS_RE.test(v) || SVG_REFUSE_DATA_HTML_RE.test(v))) return tag;
      strippedHref = true;
      return tag.replace(m[0], '');
    });
    if (strippedHref && !removed.includes('dangerous-href')) removed.push('dangerous-href');
  }
  return { cleaned: out, removed };
}

// ── External relationship targets (allowlist + audit trail) ─────────────────
// External rels (`TargetMode=External`: hyperlinks, composed links) are
// first-class output aimed at renderers that follow them, so targets are
// classified against a scheme allowlist AND recorded in a bounded in-memory
// audit trail (no console logging: library code stays silent; tests and
// operators drain the trail via getExternalRelAuditTrail).

/** Schemes an external relationship target may use (mirrors compose link policy). */
export const ALLOWED_EXTERNAL_REL_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'ftp:', 'tel:', 'sms:']);

export interface ExternalRelDecision {
  /** Raw target as classified (truncated for the trail, never echoed in errors). */
  target: string;
  /** Lowercased scheme with colon (`http:`), or '' for relative/schemeless targets. */
  scheme: string;
  /** True when the target may be emitted as an external rel. */
  allowed: boolean;
  /** Stable reason code for the trail (`allowlisted-scheme`, `relative-target`, `blocked-scheme`). */
  reason: string;
}

/**
 * Pure classifier: allowlisted absolute schemes and relative targets pass;
 * any other scheme (javascript:, data:, file:, vbscript:, ...) is blocked.
 */
export function classifyExternalRelTarget(target: unknown): ExternalRelDecision {
  const raw = typeof target === 'string' ? target : String(target ?? '');
  const trimmed = raw.trim();
  const shown = trimmed.slice(0, 128);
  const schemeMatch = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(trimmed);
  if (!schemeMatch) return { target: shown, scheme: '', allowed: true, reason: 'relative-target' };
  const scheme = `${schemeMatch[1].toLowerCase()}:`;
  if (ALLOWED_EXTERNAL_REL_SCHEMES.has(scheme)) {
    return { target: shown, scheme, allowed: true, reason: 'allowlisted-scheme' };
  }
  return { target: shown, scheme, allowed: false, reason: 'blocked-scheme' };
}

/** One audit-trail entry (bounded ring, oldest dropped). */
export interface ExternalRelAuditEntry extends ExternalRelDecision {
  /** Wall-clock millis when the decision was recorded. */
  at: number;
}

const EXTERNAL_REL_AUDIT_MAX = 500;
const externalRelAudit: ExternalRelAuditEntry[] = [];

/** Records an external-rel decision on the audit trail (never throws). */
export function auditExternalRelTarget(target: unknown): ExternalRelDecision {
  try {
    const decision = classifyExternalRelTarget(target);
    externalRelAudit.push({ ...decision, at: Date.now() });
    while (externalRelAudit.length > EXTERNAL_REL_AUDIT_MAX) externalRelAudit.shift();
    return decision;
  } catch {
    return { target: '', scheme: '', allowed: false, reason: 'blocked-scheme' };
  }
}

/** Drains a copy of the external-rel audit trail (oldest first). */
export function getExternalRelAuditTrail(): ExternalRelAuditEntry[] {
  return [...externalRelAudit];
}

/** Test hook: clears the external-rel audit trail. */
export function clearExternalRelAuditTrail(): void {
  externalRelAudit.length = 0;
}
