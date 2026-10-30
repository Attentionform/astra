import { createWordElement } from '../ooxml/ns';
import { DocumentAST, Node } from '../ast';
import { BaseOperation, createTextElement, validateVersion, findNode } from './base';
import { RPR_CHILD_ORDER, PPR_CHILD_ORDER, appendInSchemaOrder, sortChildrenBySchema } from '../ooxml/schema-order';
import { parseFontSize, requireTwips } from '../operation-schema';
import { toError } from '../errors';

export interface StyleProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  doubleStrikethrough?: boolean;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
  alignment?: 'left' | 'center' | 'right' | 'both' | 'justify';
  color?: string;       // hex without '#', e.g. '000000', 'FF0000'
  fontSize?: number | string;    // half-points (w:sz units, ECMA-376 file truth); points only via "Npt" strings
  sz?: number | string;          // half-points (w:sz units); points only via "Npt" strings
  fontFamily?: string;  // e.g. 'Arial', 'Times New Roman'
  highlight?: string;   // OOXML highlight colors: 'yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black', 'none'
  lineSpacing?: number; // in 240ths of a line: 240 = 1.0x, 276 = 1.15x, 360 = 1.5x, 480 = 2.0x
  spacingBefore?: number | string; // in DXA (20 DXA = 1pt, 240 DXA = 12pt)
  spacingAfter?: number | string;  // in DXA (20 DXA = 1pt, 120 DXA = 6pt)
  indentLeft?: number | string;    // in DXA
  indentRight?: number | string;   // in DXA
  indentFirstLine?: number | string; // in DXA
  indentHanging?: number | string;   // in DXA
  borderBottom?: { style?: string; sz?: number; color?: string };
  borderTop?: { style?: string; sz?: number; color?: string };
  borderLeft?: { style?: string; sz?: number; color?: string };
  borderRight?: { style?: string; sz?: number; color?: string };
  borderBox?: { style?: string; sz?: number; color?: string };
  // Word-level span gate: when set, run props apply only to the matched
  // substring (via in-place run splitting); paragraph props still apply
  // paragraph-wide. Omitted = legacy whole-node styling.
  match?: string;
  matchAll?: boolean;   // default false (first occurrence only)
  matchCase?: boolean;  // default true (case-sensitive)
}

export interface StyleOperation extends BaseOperation, StyleProperties {
  action: 'style';
  target: string | string[];
  properties?: StyleProperties;
  params?: StyleProperties;
}

/**
 * Per-run property applier shared by the legacy whole-node path and the
 * word-level span path. Statement-for-statement identical to the former
 * inline loop body, so match-omitted output bytes cannot drift.
 */
function applyRunProps(doc: Document, run: Element, props: StyleProperties): void {
  let rPr = run.getElementsByTagName('w:rPr')[0];
  if (!rPr) {
    rPr = createWordElement(doc, 'w:rPr');
    run.insertBefore(rPr, run.firstChild);
  }

  const applyBoolProp = (tagName: string, value?: boolean) => {
    if (value === undefined) return;
    const existing = rPr.getElementsByTagName(tagName)[0];
    if (value && !existing) {
      rPr.appendChild(createWordElement(doc, tagName));
    } else if (!value && existing) {
      rPr.removeChild(existing);
    }
  };

  applyBoolProp('w:b', props.bold);
  applyBoolProp('w:i', props.italic);
  applyBoolProp('w:u', props.underline);
  applyBoolProp('w:strike', props.strikethrough);
  applyBoolProp('w:dstrike', props.doubleStrikethrough);

  // Vertical Alignment (Superscript / Subscript)
  if (props.vertAlign !== undefined) {
    let vaEl = rPr.getElementsByTagName('w:vertAlign')[0];
    if (props.vertAlign === 'baseline') {
      if (vaEl) rPr.removeChild(vaEl);
    } else {
      if (!vaEl) {
        vaEl = createWordElement(doc, 'w:vertAlign');
        rPr.appendChild(vaEl);
      }
      vaEl.setAttribute('w:val', props.vertAlign);
    }
  }

  // Color
  if (props.color !== undefined && props.color !== null) {
    let colorEl = rPr.getElementsByTagName('w:color')[0];
    if (!colorEl) {
      colorEl = createWordElement(doc, 'w:color');
      rPr.appendChild(colorEl);
    }
    const hex = String(props.color).replace(/^#/, '').trim();
    if (hex) {
      colorEl.setAttribute('w:val', hex);
    }
  }

  // Font Size: bare numbers are half-points (w:sz units, ECMA-376 file
  // truth) on every path; points only via explicit "Npt" strings
  // (parseFontSize converts x2). Matches the replace passthrough.
  const rawSz = props.sz !== undefined ? props.sz : props.fontSize;
  if (rawSz !== undefined && rawSz !== null) {
    const hps = parseFontSize(rawSz);
    if (hps !== undefined) {
      let szEl = rPr.getElementsByTagName('w:sz')[0];
      if (!szEl) {
        szEl = createWordElement(doc, 'w:sz');
        rPr.appendChild(szEl);
      }
      szEl.setAttribute('w:val', String(hps));

      let szCsEl = rPr.getElementsByTagName('w:szCs')[0];
      if (!szCsEl) {
        szCsEl = createWordElement(doc, 'w:szCs');
        rPr.appendChild(szCsEl);
      }
      szCsEl.setAttribute('w:val', String(hps));
    }
  }

  // Font Family
  if (props.fontFamily !== undefined) {
    let rFonts = rPr.getElementsByTagName('w:rFonts')[0];
    if (!rFonts) {
      rFonts = createWordElement(doc, 'w:rFonts');
      rPr.appendChild(rFonts);
    }
    rFonts.setAttribute('w:ascii', props.fontFamily);
    rFonts.setAttribute('w:hAnsi', props.fontFamily);
  }

  // Highlight
  if (props.highlight !== undefined) {
    let hlEl = rPr.getElementsByTagName('w:highlight')[0];
    if (props.highlight === 'none') {
      if (hlEl) rPr.removeChild(hlEl);
    } else {
      if (!hlEl) {
        hlEl = createWordElement(doc, 'w:highlight');
        rPr.appendChild(hlEl);
      }
      hlEl.setAttribute('w:val', props.highlight);
    }
  }

  sortChildrenBySchema(rPr, RPR_CHILD_ORDER);
}

/**
 * Resolves match/matchAll/matchCase flat-wins: op.* ?? properties.* ??
 * params.*. A params-nested match must never silently degrade to
 * whole-paragraph styling, and a flat match always wins on conflict.
 */
function resolveMatchOpts(op: StyleOperation): { match: string | undefined; matchAll: boolean; matchCase: boolean } {
  const flat = op as unknown as Record<string, unknown>;
  const pick = (key: string): unknown => {
    const f = flat[key];
    if (f !== undefined) return f;
    for (const bag of [op.properties, op.params]) {
      if (bag && typeof bag === 'object') {
        const v = (bag as unknown as Record<string, unknown>)[key];
        if (v !== undefined) return v;
      }
    }
    return undefined;
  };
  const rawMatch = pick('match');
  let match: string | undefined;
  if (rawMatch !== undefined) {
    if (typeof rawMatch !== 'string' || rawMatch.length < 1 || rawMatch.length > 1000) {
      throw toError(
        'E_VALIDATION',
        'style \'match\' must be a non-empty string of 1-1000 characters',
        'pass match as the exact substring to style (e.g. { action: \'style\', target: \'<paraId>\', match: \'Riverbend\', bold: true }).'
      );
    }
    match = rawMatch;
  }
  return { match, matchAll: pick('matchAll') === true, matchCase: pick('matchCase') !== false };
}

/**
 * Style keys the applier and the match resolver read: flat on the op, inside
 * `properties`, or inside `params` (all three merge per key). Border entries
 * are style keys too — their config objects carry `sz`/`color` and must not
 * be mistaken for unread nesting.
 */
const READ_STYLE_KEYS = new Set<string>([
  'bold', 'italic', 'underline', 'strikethrough', 'doubleStrikethrough',
  'vertAlign', 'color', 'fontSize', 'sz', 'fontFamily', 'highlight',
  'alignment', 'lineSpacing', 'spacingBefore', 'spacingAfter',
  'indentLeft', 'indentRight', 'indentFirstLine', 'indentHanging',
  'borderBottom', 'borderTop', 'borderLeft', 'borderRight', 'borderBox',
  'match', 'matchAll', 'matchCase',
]);

/** Keys that select text but change nothing on their own. */
const SELECTOR_ONLY_KEYS = new Set<string>(['match', 'matchAll', 'matchCase']);

const BORDER_KEYS = ['borderBottom', 'borderTop', 'borderLeft', 'borderRight', 'borderBox'] as const;
const BORDER_CONFIG_KEYS = new Set<string>(['style', 'sz', 'color']);

/** First style key the op scan finds where the style reader never applies it. */
interface UnreadStyleHit {
  /** Dotted path with array indices, e.g. `params.wrapper[0]`. */
  path: string;
  /** `nested` = style key outside the read locations; `null` = null-valued style key at one; `border` = key a border config never reads. */
  kind: 'nested' | 'null' | 'border';
}

/** First key inside a border config (`borderBox: { ... }`) that the applier never reads. */
function findUnreadBorderConfigKey(owner: unknown, prefix: string): UnreadStyleHit | null {
  if (!owner || typeof owner !== 'object' || Array.isArray(owner)) return null;
  const record = owner as Record<string, unknown>;
  for (const border of BORDER_KEYS) {
    const config = record[border];
    if (!config || typeof config !== 'object' || Array.isArray(config)) continue;
    for (const key of Object.keys(config as Record<string, unknown>)) {
      if (!BORDER_CONFIG_KEYS.has(key)) return { path: `${prefix ? `${prefix}.` : ''}${border}.${key}`, kind: 'border' };
    }
  }
  return null;
}

/**
 * Path of the first style key parked where the style op never reads it, or
 * null: a style key anywhere outside the op's flat fields (`properties`/
 * `params` are read there, so their keys are skipped at the op level), inside
 * an object or array nested below any other key — arrays are containers, so
 * a wrapped list reports its indexed path (`params.wrapper[0]`) — or an
 * explicit null value on a style key at a read location (`match` keeps its
 * own non-empty-string refusal). Recognized style keys are never descended
 * into, so border configs stay exempt. Cycle-safe via a visited set and
 * uncapped: ops that reach executeStyle through the engine API are
 * JSON-serializable (measureBatchBytes rejects circular payloads before
 * routing), and a direct cyclic call still terminates.
 */
function findUnreadStyleKeyPath(owner: unknown, prefix: string, skipBagKeys: boolean, insideUnread = false, visited = new Set<object>()): UnreadStyleHit | null {
  if (!owner || typeof owner !== 'object' || visited.has(owner as object)) return null;
  visited.add(owner as object);
  if (Array.isArray(owner)) {
    for (let i = 0; i < owner.length; i++) {
      const hit = findUnreadStyleKeyPath(owner[i], `${prefix}[${i}]`, false, insideUnread, visited);
      if (hit) return hit;
    }
    return null;
  }
  const record = owner as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    const value = record[key];
    const here = prefix ? `${prefix}.${key}` : key;
    if (READ_STYLE_KEYS.has(key)) {
      if (insideUnread) return { path: here, kind: 'nested' };
      if (value === null && key !== 'match') return { path: here, kind: 'null' };
      continue;
    }
    if (skipBagKeys && (key === 'properties' || key === 'params')) continue;
    if (!value || typeof value !== 'object') continue;
    if (Array.isArray(value)) {
      const hit = findUnreadStyleKeyPath(value, here, false, true, visited);
      if (hit) return hit;
      continue;
    }
    if (Object.keys(value as Record<string, unknown>).some((k) => READ_STYLE_KEYS.has(k))) return { path: here, kind: 'nested' };
    const deeper = findUnreadStyleKeyPath(value, here, false, true, visited);
    if (deeper) return deeper;
  }
  return null;
}

function unreadStyleKeyError(hit: UnreadStyleHit) {
  const message = hit.kind === 'null'
    ? `style op sets style key "${hit.path}" to null, which applies nothing and would be silently ignored.`
    : hit.kind === 'border'
      ? `style op sets "${hit.path}", but a border config only reads style, sz and color — the key would be silently ignored.`
      : `style op nests style keys under "${hit.path}", which the engine never reads — nested keys would be silently ignored.`;
  return toError(
    'E_VALIDATION',
    message,
    'Pass style keys flat on the op, inside "properties", or inside "params" (all three merge per key), e.g. ' +
      '{ action: \'style\', target: \'<paraId>\', color: \'FF0000\', match: \'Riverbend\' }; pass real values (omit the key instead of null) and retry.'
  );
}

interface MappedRun {
  run: Element;
  text: string;
  start: number;
  end: number;
}

interface MatchSpan {
  start: number;
  end: number;
}

/**
 * Visible text of one run: w:t concatenation with tabs/breaks counted as one
 * character each so match offsets stay aligned with the run-boundary map.
 */
function runVisibleText(run: Element): string {
  let out = '';
  const kids = run.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const c = kids[i] as unknown as Element;
    if (!c || c.nodeType !== 1) continue;
    if (c.tagName === 'w:t') out += c.textContent || '';
    else if (c.tagName === 'w:tab') out += '\t';
    else if (c.tagName === 'w:br') out += '\n';
  }
  return out;
}

function buildRunMap(pEl: Element): { runs: MappedRun[]; text: string } {
  const els = Array.from(pEl.getElementsByTagName('w:r')) as unknown as Element[];
  const runs: MappedRun[] = [];
  let text = '';
  for (const run of els) {
    const t = runVisibleText(run);
    runs.push({ run, text: t, start: text.length, end: text.length + t.length });
    text += t;
  }
  return { runs, text };
}

export function findMatchSpans(text: string, match: string, matchCase: boolean, matchAll: boolean): MatchSpan[] {
  // Defensive empty-needle clamp: the schema gate rejects empty matches,
  // but a direct call with matchAll would otherwise spin forever on
  // indexOf('', from) === from with zero advance.
  if (!match || match.length === 0) return [];
  const hay = matchCase ? text : text.toLowerCase();
  const needle = matchCase ? match : match.toLowerCase();
  const out: MatchSpan[] = [];
  let from = 0;
  for (;;) {
    const idx = hay.indexOf(needle, from);
    if (idx < 0) break;
    out.push({ start: idx, end: idx + needle.length });
    if (!matchAll) break;
    from = idx + needle.length;
  }
  return out;
}

function notFoundError(match: string, nodeId: string) {
  return toError(
    'E_NOT_FOUND',
    `text "${match}" not found in node "${nodeId}".`,
    'check the paragraph text via inspect_document and retry with an exact matching substring.'
  );
}

function boundaryError(match: string, nodeId: string) {
  return toError(
    'E_VALIDATION',
    `match "${match}" crosses a hyperlink or field boundary in node "${nodeId}".`,
    'pick a match substring that lies fully inside (or outside) the linked text, then retry.'
  );
}

function surrogateError(match: string, nodeId: string) {
  return toError(
    'E_VALIDATION',
    `style 'match' would split a surrogate pair in node "${nodeId}".`,
    'pick a match substring aligned to full characters (keep emoji and paired surrogates intact), then retry.'
  );
}

function isLeadSurrogate(code: number): boolean {
  return code >= 0xd800 && code <= 0xdbff;
}

function isTrailSurrogate(code: number): boolean {
  return code >= 0xdc00 && code <= 0xdfff;
}

function hasLoneSurrogate(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (isLeadSurrogate(c)) {
      if (i + 1 >= s.length || !isTrailSurrogate(s.charCodeAt(i + 1))) return true;
      i++;
    } else if (isTrailSurrogate(c)) {
      return true;
    }
  }
  return false;
}

/**
 * Refuses a match that would sever a UTF-16 surrogate pair: a lone
 * lead/trail surrogate in the needle, or any span cutting a lead/trail
 * pair apart in the haystack. Runs before any mutation on every path.
 */
function assertSurrogateSafe(text: string, match: string, spans: MatchSpan[], nodeId: string): void {
  if (hasLoneSurrogate(match)) throw surrogateError(match, nodeId);
  for (const s of spans) {
    if (s.start > 0 && isLeadSurrogate(text.charCodeAt(s.start - 1)) && isTrailSurrogate(text.charCodeAt(s.start))) {
      throw surrogateError(match, nodeId);
    }
    if (s.end < text.length && isLeadSurrogate(text.charCodeAt(s.end - 1)) && isTrailSurrogate(text.charCodeAt(s.end))) {
      throw surrogateError(match, nodeId);
    }
  }
}

/**
 * Runs overlapped by a span; fails closed when the span is absent or spans
 * runs under different containers (hyperlink/field boundary cross).
 */
function spanRuns(runs: MappedRun[], span: MatchSpan, nodeId: string, match: string): MappedRun[] {
  const touched = runs.filter((r) => r.start < span.end && r.end > span.start);
  if (touched.length === 0) throw notFoundError(match, nodeId);
  const anchor: unknown = touched[0].run.parentNode;
  if (touched.some((t) => (t.run.parentNode as unknown) !== anchor)) throw boundaryError(match, nodeId);
  return touched;
}

function directRPr(run: Element): Element | null {
  const kids = run.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const c = kids[i] as unknown as Element;
    if (c && c.nodeType === 1 && c.tagName === 'w:rPr') return c;
  }
  return null;
}

/** Fresh run shell: copied run attributes + cloned rPr, no content. */
function cloneRunShell(doc: Document, src: Element): Element {
  const nu = createWordElement(doc, 'w:r');
  const attrs = src.attributes;
  if (attrs) {
    for (let i = 0; i < attrs.length; i++) {
      const at = attrs.item(i);
      if (at) nu.setAttribute(at.name, at.value);
    }
  }
  const pr = directRPr(src);
  if (pr) nu.appendChild(pr.cloneNode(true));
  return nu;
}

/**
 * Splits a run in place at local text cuts (ascending offsets strictly
 * inside the run text). Content redistribute: wholly-inside w:t nodes move
 * untouched; straddling w:t text splits via fresh whitespace-preserving
 * w:t pieces; tabs/breaks ride atomically; w:rPr is cloned per segment,
 * never moved. Returns the segments in order with the matched one marked.
 * Sibling runs are never merged: middle-run formatting survives.
 */
function partitionRun(doc: Document, runEl: Element, cuts: number[], matchedSeg: number): { segments: Element[]; matched: Element } {
  const kids = Array.from(runEl.childNodes);
  const lens = kids.map((k) => {
    const el = k as unknown as Element;
    if (k.nodeType !== 1) return 0;
    if (el.tagName === 'w:t') return (el.textContent || '').length;
    if (el.tagName === 'w:tab' || el.tagName === 'w:br') return 1;
    return 0;
  });
  const total = lens.reduce((a, b) => a + b, 0);
  const edges = [0, ...cuts.filter((c) => c > 0 && c < total).sort((a, b) => a - b), total]
    .filter((b, i, arr) => i === 0 || b !== arr[i - 1]);
  const buckets: ChildNode[][] = edges.slice(0, -1).map(() => []);
  const segOf = (pos: number): number => {
    let s = 0;
    for (let i = 1; i < edges.length - 1; i++) {
      if (edges[i] <= pos) s = i;
    }
    return s;
  };
  let cum = 0;
  for (let k = 0; k < kids.length; k++) {
    const node = kids[k] as unknown as Element;
    if (node.nodeType === 1 && node.tagName === 'w:rPr') continue;
    const len = lens[k];
    if (node.nodeType === 1 && node.tagName === 'w:t' && len > 0) {
      let off = 0;
      while (off < len) {
        const seg = segOf(cum);
        const take = Math.min(len - off, edges[seg + 1] - cum);
        const piece = (node.textContent || '').substring(off, off + take);
        if (piece.length > 0) buckets[seg].push(createTextElement(doc, piece));
        off += take;
        cum += take;
      }
      continue;
    }
    buckets[segOf(cum)].push(kids[k]);
    cum += len;
  }
  const parent = runEl.parentNode;
  if (!parent || buckets.every((b) => b.length === 0)) return { segments: [runEl], matched: runEl };
  const segments: Element[] = [];
  for (const bucket of buckets) {
    if (bucket.length === 0) continue;
    const shell = cloneRunShell(doc, runEl);
    for (const n of bucket) shell.appendChild(n);
    segments.push(shell);
  }
  if (segments.length === 0) return { segments: [runEl], matched: runEl };
  for (const s of segments) parent.insertBefore(s, runEl);
  parent.removeChild(runEl);
  const clamped = Math.min(Math.max(matchedSeg, 0), segments.length - 1);
  return { segments, matched: segments[clamped] };
}

function applySingleSpan(
  doc: Document,
  runs: MappedRun[],
  span: MatchSpan,
  props: StyleProperties,
  nodeId: string,
  match: string
): void {
  const touched = spanRuns(runs, span, nodeId, match);
  if (touched.length === 1) {
    const t = touched[0];
    const a = span.start - t.start;
    const b = span.end - t.start;
    if (a <= 0 && b >= t.text.length) {
      applyRunProps(doc, t.run, props);
      return;
    }
    const cuts: number[] = [];
    if (a > 0) cuts.push(a);
    if (b < t.text.length) cuts.push(b);
    const { matched } = partitionRun(doc, t.run, cuts, a > 0 ? 1 : 0);
    applyRunProps(doc, matched, props);
    return;
  }
  const first = touched[0];
  const last = touched[touched.length - 1];
  const headCut = span.start - first.start;
  if (headCut <= 0) {
    applyRunProps(doc, first.run, props);
  } else {
    const { matched } = partitionRun(doc, first.run, [headCut], 1);
    applyRunProps(doc, matched, props);
  }
  for (let k = 1; k < touched.length - 1; k++) {
    applyRunProps(doc, touched[k].run, props);
  }
  const tailCut = span.end - last.start;
  if (tailCut >= last.text.length) {
    applyRunProps(doc, last.run, props);
  } else {
    const { matched } = partitionRun(doc, last.run, [tailCut], 0);
    applyRunProps(doc, matched, props);
  }
}

/**
 * Word-level styling: paragraph props ride paragraph-wide (applied by the
 * caller); run props land on the matched span(s) only via in-place run
 * splitting. Cross-container matches refuse before any mutation; spans apply
 * last-first so earlier offsets stay valid.
 */
function applySpanStyle(
  doc: Document,
  pEl: Element,
  nodeId: string,
  props: StyleProperties,
  match: string,
  matchAll: boolean,
  matchCase: boolean,
  hasRunProps: boolean
): void {
  if (!pEl || (pEl as Element).tagName !== 'w:p') {
    throw toError(
      'E_VALIDATION',
      `style 'match' only supports paragraph nodes ('w:p'), got '${(pEl as Element)?.tagName || 'unknown'}' for node "${nodeId}".`,
      `Target a paragraph node id from inspect (e.g. { action: 'style', target: '<paraId>', match: '...' }); for text in a table cell, target '<tableId>:r<row>c<col>' (rows and columns count from 1).`
    );
  }
  const first = buildRunMap(pEl);
  const spans = findMatchSpans(first.text, match, matchCase, matchAll);
  if (spans.length === 0) throw notFoundError(match, nodeId);
  assertSurrogateSafe(first.text, match, spans, nodeId);
  for (const s of spans) spanRuns(first.runs, s, nodeId, match);
  if (!hasRunProps) return;
  const ordered = [...spans].sort((a, b) => b.start - a.start);
  for (const s of ordered) {
    applySingleSpan(doc, buildRunMap(pEl).runs, s, props, nodeId, match);
  }
}

export function executeStyle(op: StyleOperation, ast: DocumentAST): void {
  const targets = Array.isArray(op.target) ? op.target : [op.target];
  if (targets.length === 0 || !targets[0]) throw new Error("Style operation requires a target");

  // Silent-ignore gate: style keys are read from the op's flat fields,
  // `properties`, and `params` only. Keys parked anywhere else (e.g. a
  // params-nested params bag) would be quietly dropped, so refuse before any
  // target is resolved or mutated.
  const unread = findUnreadStyleKeyPath(op, '', true) ??
    findUnreadStyleKeyPath(op.properties, 'properties', false) ??
    findUnreadStyleKeyPath(op.params, 'params', false);
  if (unread) throw unreadStyleKeyError(unread);
  const unreadBorder = findUnreadBorderConfigKey(op, '') ??
    findUnreadBorderConfigKey(op.properties, 'properties') ??
    findUnreadBorderConfigKey(op.params, 'params');
  if (unreadBorder) throw unreadStyleKeyError(unreadBorder);

  // Per-key flat-wins merge (params < properties < flat), mirroring
  // resolveMatchOpts: mixed bags merge instead of one whole bag winning.
  const props: StyleProperties = {
    ...(op.params || {}),
    ...(op.properties || {}),
    ...((op as unknown as StyleProperties) || {}),
  };
  const matchOpts = resolveMatchOpts(op);

  // Nothing-to-apply gate: match/matchAll/matchCase only select text, so an op
  // with no real style key would report success while changing nothing.
  const sentinel = props as unknown as Record<string, unknown>;
  const appliesSomething = [...READ_STYLE_KEYS].some(
    (key) => !SELECTOR_ONLY_KEYS.has(key) && sentinel[key] !== undefined && sentinel[key] !== null
  );
  if (!appliesSomething) {
    throw toError(
      'E_VALIDATION',
      'style op sets no style property, so nothing would change.',
      "Pass at least one style key (for example bold, color, fontSize or alignment) flat on the op, inside \"properties\", or inside \"params\"; match, matchAll and matchCase only select text."
    );
  }

  // Run-props presence gate (shared by pre-validation and the apply pass).
  const hasRunProps = props.bold !== undefined ||
    props.italic !== undefined ||
    props.underline !== undefined ||
    props.strikethrough !== undefined ||
    props.doubleStrikethrough !== undefined ||
    props.vertAlign !== undefined ||
    props.color !== undefined ||
    props.fontSize !== undefined ||
    props.sz !== undefined ||
    props.fontFamily !== undefined ||
    props.highlight !== undefined;

  // KILL-1 atomicity gate: when `match` is set, every target's spans
  // (presence + hyperlink/field boundary checks) validate BEFORE any pPr
  // or run mutation on any target, so a miss/boundary/non-paragraph
  // refusal leaves zero bytes touched. Success paths re-enter the same
  // checks inside applySpanStyle and behave identically.
  const resolvedNodes: Node[] = [];
  for (const target of targets) {
    const targetNode: Node | undefined = findNode(ast, target);

    if (!targetNode) throw new Error(`Node ${target} not found`);
    validateVersion(targetNode, op.expectedVersion);
    resolvedNodes.push(targetNode);

    if (matchOpts.match !== undefined) {
      const doc = targetNode.xml.ownerDocument;
      if (!doc) continue;
      const pEl = targetNode.xml as unknown as Element;
      if (!pEl || (pEl as Element).tagName !== 'w:p') {
        throw toError(
          'E_VALIDATION',
          `style 'match' only supports paragraph nodes ('w:p'), got '${(pEl as Element)?.tagName || 'unknown'}' for node "${target}".`,
          `Target a paragraph node id from inspect (e.g. { action: 'style', target: '<paraId>', match: '...' }); for text in a table cell, target '<tableId>:r<row>c<col>' (rows and columns count from 1).`
        );
      }
      const map = buildRunMap(pEl);
      const spans = findMatchSpans(map.text, matchOpts.match, matchOpts.matchCase, matchOpts.matchAll);
      if (spans.length === 0) throw notFoundError(matchOpts.match, target);
      assertSurrogateSafe(map.text, matchOpts.match, spans, target);
      for (const s of spans) spanRuns(map.runs, s, target, matchOpts.match);
    }
  }

  for (const targetNode of resolvedNodes) {

    const doc = targetNode.xml.ownerDocument;
    if (!doc) continue;

    // Apply Paragraph Properties (Alignment, Spacing, Indentation, Borders)
    const hasBorders = props.borderBottom !== undefined ||
      props.borderTop !== undefined ||
      props.borderLeft !== undefined ||
      props.borderRight !== undefined ||
      props.borderBox !== undefined;

    const hasPPr = props.alignment ||
      props.lineSpacing !== undefined ||
      props.spacingBefore !== undefined ||
      props.spacingAfter !== undefined ||
      props.indentLeft !== undefined ||
      props.indentRight !== undefined ||
      props.indentFirstLine !== undefined ||
      props.indentHanging !== undefined ||
      hasBorders;

    if (hasPPr) {
      let pPr = targetNode.xml.getElementsByTagName('w:pPr')[0];
      if (!pPr) {
        pPr = createWordElement(doc, 'w:pPr');
        targetNode.xml.insertBefore(pPr, targetNode.xml.firstChild);
      }

      if (props.alignment) {
        let jc = pPr.getElementsByTagName('w:jc')[0];
        if (!jc) {
          jc = createWordElement(doc, 'w:jc');
          appendInSchemaOrder(pPr, jc, PPR_CHILD_ORDER);
        }
        let alignVal = props.alignment;
        if ((alignVal as string) === 'justify') alignVal = 'both';
        jc.setAttribute('w:val', alignVal);
      }

      // Paragraph Spacing (<w:spacing w:line="..." w:lineRule="auto" w:before="..." w:after="..."/>)
      if (props.lineSpacing !== undefined || props.spacingBefore !== undefined || props.spacingAfter !== undefined) {
        let sp = pPr.getElementsByTagName('w:spacing')[0];
        if (!sp) {
          sp = createWordElement(doc, 'w:spacing');
          appendInSchemaOrder(pPr, sp, PPR_CHILD_ORDER);
        }
        if (props.lineSpacing !== undefined) {
          sp.setAttribute('w:line', String(props.lineSpacing));
          sp.setAttribute('w:lineRule', 'auto');
        }
        if (props.spacingBefore !== undefined) {
          sp.setAttribute('w:before', String(requireTwips('spacingBefore', props.spacingBefore)));
        }
        if (props.spacingAfter !== undefined) {
          sp.setAttribute('w:after', String(requireTwips('spacingAfter', props.spacingAfter)));
        }
      }

      // Paragraph Indentation (<w:ind w:left="..." w:right="..." w:firstLine="..." w:hanging="..."/>)
      if (props.indentLeft !== undefined || props.indentRight !== undefined || props.indentFirstLine !== undefined || props.indentHanging !== undefined) {
        let ind = pPr.getElementsByTagName('w:ind')[0];
        if (!ind) {
          ind = createWordElement(doc, 'w:ind');
          appendInSchemaOrder(pPr, ind, PPR_CHILD_ORDER);
        }
        if (props.indentLeft !== undefined) ind.setAttribute('w:left', String(requireTwips('indentLeft', props.indentLeft)));
        if (props.indentRight !== undefined) ind.setAttribute('w:right', String(requireTwips('indentRight', props.indentRight)));
        if (props.indentFirstLine !== undefined) ind.setAttribute('w:firstLine', String(requireTwips('indentFirstLine', props.indentFirstLine)));
        if (props.indentHanging !== undefined) ind.setAttribute('w:hanging', String(requireTwips('indentHanging', props.indentHanging)));
      }

      // Paragraph Borders (<w:pBdr><w:bottom .../><w:top .../></w:pBdr>)
      if (hasBorders) {
        let pBdr = pPr.getElementsByTagName('w:pBdr')[0];
        if (!pBdr) {
          pBdr = createWordElement(doc, 'w:pBdr');
          appendInSchemaOrder(pPr, pBdr, PPR_CHILD_ORDER);
        }

        const applyBdr = (side: string, config?: { style?: string; sz?: number; color?: string }) => {
          if (!config) return;
          let el = pBdr.getElementsByTagName(`w:${side}`)[0];
          if (!el) {
            el = createWordElement(doc, `w:${side}`);
            pBdr.appendChild(el);
          }
          el.setAttribute('w:val', config.style || 'single');
          el.setAttribute('w:sz', String(config.sz ?? 12));
          el.setAttribute('w:space', '1');
          el.setAttribute('w:color', (config.color || 'CBD5E1').replace(/^#/, ''));
        };

        if (props.borderBox) {
          applyBdr('top', props.borderBox);
          applyBdr('bottom', props.borderBox);
          applyBdr('left', props.borderBox);
          applyBdr('right', props.borderBox);
        }
        applyBdr('bottom', props.borderBottom);
        applyBdr('top', props.borderTop);
        applyBdr('left', props.borderLeft);
        applyBdr('right', props.borderRight);
      }

      sortChildrenBySchema(pPr, PPR_CHILD_ORDER);
    }

    // Apply Run Properties (Bold, Italic, Underline, Strikethrough, Color, FontSize, FontFamily, Highlight, VertAlign)
    if (matchOpts.match !== undefined) {
      applySpanStyle(doc, targetNode.xml, targetNode.id, props, matchOpts.match, matchOpts.matchAll, matchOpts.matchCase, hasRunProps);
    } else if (hasRunProps) {
      const runs = targetNode.xml.getElementsByTagName('w:r');
      for (let i = 0; i < runs.length; i++) {
        applyRunProps(doc, runs[i] as unknown as Element, props);
      }
    }

    targetNode.version += 1;
  }
}

