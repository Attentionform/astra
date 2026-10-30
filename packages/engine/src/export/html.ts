import { DocumentAST } from '../ast';
import type { DocxParser } from '../parser';
import { DOMParser } from '@xmldom/xmldom';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { ommlToMathmlDisplay } from '../ooxml/omml';
import { isMonospaceRun, hasListNumbering, isOrderedListItem, codeBlockLanguage } from './markdown';

/** Narrow export duck-typing without behavior change (replaces `any`). */
export type ExportNodeLike = {
  id?: string;
  type?: string;
  caption?: string;
  captionPosition?: string;
  headers?: string[] | unknown;
  rows?: unknown;
  xml?: unknown;
  metadata?: unknown;
};

interface FootnoteZipCapable {
  getZipAdapter?: () => { readTextOrNull: (p: string) => string | null };
  zip?: { getEntry: (p: string) => { getData: () => Buffer } | null | undefined };
}

function asFootnoteZipCapable(parser: unknown): FootnoteZipCapable {
  if (typeof parser !== 'object' || parser === null) return {};
  const rec = parser as Record<string, unknown>;
  const out: FootnoteZipCapable = {};
  if (typeof rec['getZipAdapter'] === 'function') {
    const fn = rec['getZipAdapter'] as unknown as () => { readTextOrNull: (p: string) => string | null };
    out.getZipAdapter = fn.bind(parser);
  }
  if (rec['zip'] !== null && typeof rec['zip'] === 'object') {
    const zipRec = rec['zip'] as Record<string, unknown>;
    if (typeof zipRec['getEntry'] === 'function') {
      const getEntry = zipRec['getEntry'] as unknown as (p: string) => { getData: () => Buffer } | null | undefined;
      out.zip = { getEntry: getEntry.bind(rec['zip']) };
    }
  }
  return out;
}

function metaTextOf(node: ExportNodeLike): string {
  const rec = node.metadata as unknown as Record<string, unknown> | undefined;
  const t: unknown = rec?.['text'];
  if (typeof t === 'string') return t;
  if (t) return t as unknown as string;
  return '';
}

function metaRecordOf(node: ExportNodeLike): Record<string, unknown> {
  if (typeof node.metadata === 'object' && node.metadata !== null) {
    return node.metadata as unknown as Record<string, unknown>;
  }
  return {};
}

/** E4: display-equation fallback — MathML when convertible, else the legacy placeholder. Never throws. */
function equationDisplayHtml(xml: Element | null | undefined, latexFallback: string): string {
  try {
    const mathml = ommlToMathmlDisplay(xml);
    if (mathml !== '') return mathml;
  } catch { /* fall through to the loud placeholder */ }
  const label = latexFallback ? `[Equation: ${escapeHtml(latexFallback)}]` : '[Equation]';
  return `<span class="equation">${label}</span>`;
}

/** OP-23: resolve embedded image bytes for a node via a:blip r:embed -> relationships -> mediaFiles. */
function resolveNodeImage(ast: DocumentAST, node: ExportNodeLike): { base64: string; mime: string; alt: string } | null {
  try {
    const xml = node?.xml as Element | undefined;
    let embedId: string | null = null;
    let alt = metaTextOf(node).trim() || 'Image';
    if (xml && typeof (xml as unknown as Element).getElementsByTagName === 'function') {
      const blips = (xml as Element).getElementsByTagName('a:blip');
      for (let i = 0; i < blips.length; i++) {
        const id = (blips[i] as Element).getAttribute('r:embed');
        if (id) { embedId = id; break; }
      }
      const docPrs = (xml as Element).getElementsByTagName('wp:docPr');
      if (docPrs.length > 0) {
        const descr = (docPrs[0] as Element).getAttribute('descr') || (docPrs[0] as Element).getAttribute('title');
        if (descr) alt = descr;
      }
    }
    if (!embedId) return null;
    const rel = (ast.relationships as Record<string, { target?: string }>)?.[embedId];
    const target: string = rel?.target || '';
    const base = target.split('/').pop() || '';
    const mediaFiles = (ast.mediaFiles || {}) as Record<string, Buffer | Uint8Array>;
    const buf: Buffer | Uint8Array | undefined =
      (base && mediaFiles[base]) || (target && (mediaFiles[target] || mediaFiles[`word/${target}`])) || undefined;
    if (!buf) return null;
    const data = Buffer.isBuffer(buf) ? buf : Buffer.from(buf as Uint8Array);
    const ext = (base.split('.').pop() || 'png').toLowerCase();
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
      : ext === 'gif' ? 'image/gif'
      : ext === 'webp' ? 'image/webp'
      : ext === 'bmp' ? 'image/bmp'
      : ext === 'svg' ? 'image/svg+xml'
      : 'image/png';
    return { base64: data.toString('base64'), mime, alt };
  } catch {
    return null;
  }
}

function getParagraphAlignment(xmlNode: Element | null | undefined): string | null {
  if (!xmlNode || typeof (xmlNode as unknown as Element).getElementsByTagName !== 'function') return null;
  const pPr = xmlNode.getElementsByTagName('w:pPr')[0];
  if (!pPr) return null;
  const jc = pPr.getElementsByTagName('w:jc')[0];
  if (!jc) return null;
  const val = (jc.getAttribute('w:val') || '').toLowerCase();
  if (val === 'center') return 'center';
  if (val === 'right' || val === 'end') return 'right';
  if (val === 'both' || val === 'justify' || val === 'distribute') return 'justify';
  if (val === 'left' || val === 'start') return 'left';
  return null;
}

function isDivider(node: ExportNodeLike, text?: string): boolean {
  const type = String(node.type || '').toLowerCase();
  if (type === 'divider' || type === 'horizontalrule' || type === 'hr') {
    return true;
  }
  // Code blocks carry a box border (w:pBdr) by design — they are never
  // dividers, even though the bottom edge matches the rule below.
  if (isCodeBlockParagraph(node)) {
    return false;
  }
  const trimmed = (text ?? metaTextOf(node) ?? '').trim();
  if (trimmed === '---' || trimmed === '***' || trimmed === '___') {
    return true;
  }
  if ((metaRecordOf(node))?.['isDivider']) {
    return true;
  }
  if (node.xml && typeof (node.xml as unknown as Element).getElementsByTagName === 'function') {
    const pBdrs = (node.xml as Element).getElementsByTagName('w:pBdr');
    for (let i = 0; i < pBdrs.length; i++) {
      const bottoms = (pBdrs[i] as Element).getElementsByTagName('w:bottom');
      if (bottoms.length > 0) {
        const val = (bottoms[0] as Element).getAttribute('w:val');
        if (!val || (val !== 'none' && val !== 'nil')) {
          return true;
        }
      }
    }
  }
  return false;
}

function renderRun(r: Element): string {
  const rPr = r.getElementsByTagName('w:rPr')[0];
  const styles: string[] = [];
  if (rPr) {
    const b = rPr.getElementsByTagName('w:b')[0];
    if (b) {
      const bVal = b.getAttribute('w:val');
      if (bVal !== '0' && bVal !== 'false' && bVal !== 'off') {
        styles.push('font-weight: bold;');
      }
    }

    const i = rPr.getElementsByTagName('w:i')[0];
    if (i) {
      const iVal = i.getAttribute('w:val');
      if (iVal !== '0' && iVal !== 'false' && iVal !== 'off') {
        styles.push('font-style: italic;');
      }
    }

    const u = rPr.getElementsByTagName('w:u')[0];
    if (u) {
      const uVal = u.getAttribute('w:val');
      if (!uVal || (uVal !== 'none' && uVal !== 'false')) {
        styles.push('text-decoration: underline;');
      }
    }

    const color = rPr.getElementsByTagName('w:color')[0];
    if (color) {
      const rawVal = color.getAttribute('w:val');
      if (rawVal && rawVal !== 'auto') {
        const val = rawVal.replace(/^#/, '');
        if (/^[0-9a-fA-F]{3,8}$/.test(val)) {
          styles.push(`color: #${val};`);
        }
      }
    }

    const sz = rPr.getElementsByTagName('w:sz')[0];
    if (sz) {
      const val = sz.getAttribute('w:val');
      if (val) {
        const pt = parseFloat(val) / 2;
        if (!isNaN(pt)) {
          styles.push(`font-size: ${pt}pt;`);
        }
      }
    }

    const rFonts = rPr.getElementsByTagName('w:rFonts')[0];
    if (rFonts) {
      const font = rFonts.getAttribute('w:ascii') || rFonts.getAttribute('w:hAnsi');
      if (font) {
        const sanitizedFont = font.replace(/['"\\;{}]/g, '').trim();
        if (sanitizedFont) {
          styles.push(`font-family: '${sanitizedFont}', sans-serif;`);
        }
      }
    }
  }

  const tEls = Array.from(r.getElementsByTagName('w:t')).concat(
    Array.from(r.getElementsByTagName('w:delText'))
  );
  let text = '';
  for (const t of tEls) {
    text += t.textContent || '';
  }

  if (!text) {
    if (r.getElementsByTagName('w:br').length > 0) return '<br/>';
    return '';
  }

  if (styles.length > 0) {
    return `<span style="${styles.join(' ')}">${escapeHtml(text)}</span>`;
  }
  return escapeHtml(text);
}

function escapeAttr(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/** Wrap tracked-change inner HTML with revision metadata from w:author /
 * w:date / w:id as title + data-* attrs. The outer del/ins tags stay
 * byte-identical so existing redline selectors keep matching. */
function wrapTrackedMeta(el: Element, inner: string, verb: 'Deleted' | 'Inserted'): string {
  const author = el.getAttribute('w:author') || el.getAttribute('w15:author') || '';
  const date = el.getAttribute('w:date') || '';
  const id = el.getAttribute('w:id') || '';
  if (!author && !date && !id) return inner;
  const bits: string[] = [];
  if (author && date) bits.push(`${verb} by ${author} on ${date}`);
  else if (author) bits.push(`${verb} by ${author}`);
  else if (date) bits.push(`${verb} on ${date}`);
  else bits.push(verb);
  const title = ` title="${escapeAttr(bits.join(''))}"`;
  const dataAuthor = author ? ` data-author="${escapeAttr(author)}"` : '';
  const dataDate = date ? ` data-date="${escapeAttr(date)}"` : '';
  const dataId = id ? ` data-id="${escapeAttr(id)}"` : '';
  return `<span${title}${dataAuthor}${dataDate}${dataId}>${inner}</span>`;
}

/**
 * OX-07 exporter drop policy (per-type: render vs placeholder vs omit vs warn).
 *
 * - footnotes/endnotes → RENDER inline `<sup>[id]</sup>` (footnotes) /
 *   `<sup>[en-id]</sup>` (endnotes) + an endnotes `<section>` with one
 *   `<li>` per ref (`fn-N` / `en-N`). Refs are namespaced so a shared
 *   numeric id keeps both entries. Bodies resolve via extractHtmlNoteBodies
 *   (ast.footnotes/ast.endnotes, then word/footnotes.xml /
 *   word/endnotes.xml through the optional parser — same `fn:`/`en:`
 *   keying as the Markdown path); the `[Footnote N: see document]`
 *   placeholder only survives when no body is reachable, never as a
 *   silent drop.
 * - comments → RENDER a `[Comment …]` badge on the covered paragraph (bodies
 *   live in word/comments.xml; the covered TEXT is always preserved).
 * - fields + TOC → RENDER display text, OMIT field codes (`w:instrText`,
 *   `w:fldChar`, `w:instr` attrs never emit).
 * - tracked changes → RENDER redline view (P1-15): `w:ins` kept (wrapped in
 *   `<ins class="tracked-ins">`), `w:del` rendered (`<del class="tracked-del">`).
 *   Both carry revision metadata (w:author/w:date/w:id) on an inner span via
 *   title + data-author/data-date/data-id attrs; the outer tags stay exact.
 * - headers/footers → RENDER deduped `<header>`/`<footer>` blocks.
 * - equations → DISPLAY (m:oMathPara / standalone equation nodes) renders
 *   native MathML `<math display="block">`; INLINE (bare m:oMath inside run
 *   flow) keeps the `[Equation: latex]` placeholder span.
 * - unknown node types → RENDER `metadata.text` when present, else an
 *   explicit `[Unsupported node type: X]` placeholder (warns loudly, drops
 *   nothing silently).
 */
interface ParagraphRender {
  html: string;
  footnoteIds: string[];
  commentIds: string[];
}

function renderXmlChildren(
  parent: Element,
  ctx: { footnoteIds: string[]; commentIds: string[] },
  ast?: DocumentAST
): string {
  let out = '';
  for (let i = 0; i < parent.childNodes.length; i++) {
    const ch = parent.childNodes[i] as Element;
    if (!ch || (ch as unknown as Element).nodeType !== 1) continue;
    out += renderParagraphChild(ch, ctx, ast);
  }
  return out;
}

function renderParagraphChild(
  el: Element,
  ctx: { footnoteIds: string[]; commentIds: string[] },
  ast?: DocumentAST
): string {
  const tag = el.tagName || '';
  if (tag === 'w:delText') {
    return escapeHtml(el.textContent || '');
  }
  if (tag === 'w:del') {
    const inner = renderXmlChildren(el, ctx, ast);
    if (!inner) return '';
    return `<del class="tracked-del">${wrapTrackedMeta(el, inner, 'Deleted')}</del>`;
  }
  if (tag === 'w:ins') {
    const inner = renderXmlChildren(el, ctx, ast);
    if (!inner) return '';
    return `<ins class="tracked-ins">${wrapTrackedMeta(el, inner, 'Inserted')}</ins>`;
  }
  if (tag === 'w:r') {
    // Runs can host inline specials (footnote refs, comment refs, OMML) —
    // renderRun only reads w:t, so surface those here (never dropped).
    let extra = '';
    const runFns = el.getElementsByTagName('w:footnoteReference');
    for (let k = 0; k < runFns.length; k++) {
      const id = (runFns[k] as Element).getAttribute('w:id') || '?';
      const key = `fn:${id}`;
      if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
      extra += `<sup class="footnote-ref">[${escapeHtml(id)}]</sup>`;
    }
    const runEnds = el.getElementsByTagName('w:endnoteReference');
    for (let k = 0; k < runEnds.length; k++) {
      const id = (runEnds[k] as Element).getAttribute('w:id') || '?';
      const key = `en:${id}`;
      if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
      extra += `<sup class="footnote-ref">[en-${escapeHtml(id)}]</sup>`;
    }
    const runCmts = el.getElementsByTagName('w:commentReference');
    for (let k = 0; k < runCmts.length; k++) {
      const id = (runCmts[k] as Element).getAttribute('w:id') || '?';
      if (!ctx.commentIds.includes(id)) ctx.commentIds.push(id);
    }
    const runMaths = el.getElementsByTagName('m:oMath');
    for (let k = 0; k < runMaths.length; k++) {
      const latex = Array.from((runMaths[k] as Element).getElementsByTagName('m:t')).map(t => (t as Element).textContent || '').join('');
      extra += `<span class="equation">${latex ? `[Equation: ${escapeHtml(latex)}]` : '[Equation]'}</span>`;
    }
    const html = renderRun(el) + extra;
    if (!html) return '';
    return html;
  }
  // Footnote/endnote reference → inline superscript marker, body in endnotes
  // section. Keys are namespaced (`fn:` vs `en:`) so a shared numeric id
  // keeps both entries instead of collapsing.
  if (tag === 'w:footnoteReference') {
    const id = el.getAttribute('w:id') || '?';
    const key = `fn:${id}`;
    if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
    return `<sup class="footnote-ref">[${escapeHtml(id)}]</sup>`;
  }
  if (tag === 'w:endnoteReference') {
    const id = el.getAttribute('w:id') || '?';
    const key = `en:${id}`;
    if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
    return `<sup class="footnote-ref">[en-${escapeHtml(id)}]</sup>`;
  }
  // Comment markers → badge rendered once per paragraph below.
  if (tag === 'w:commentRangeStart' || tag === 'w:commentRangeEnd' || tag === 'w:commentReference') {
    const id = el.getAttribute('w:id') || '';
    if (id && !ctx.commentIds.includes(id)) ctx.commentIds.push(id);
    else if (!id && !ctx.commentIds.includes('?')) ctx.commentIds.push('?');
    return '';
  }
  // Field codes carry no display text — omit; simple fields render children.
  if (tag === 'w:fldChar' || tag === 'w:instrText') return '';
  if (tag === 'w:fldSimple') return renderXmlChildren(el, ctx, ast);
  // Equation → DISPLAY (m:oMathPara) renders native MathML <math display>;
  // inline (bare m:oMath inside run flow) keeps the placeholder span.
  if (tag === 'm:oMathPara') {
    const math = equationDisplayHtml(el, '');
    return math;
  }
  if (tag === 'm:oMath') {
    const texts = Array.from(el.getElementsByTagName('m:t')).map(t => t.textContent || '').join('');
    const latex = texts || Array.from(el.getElementsByTagName('w:t')).map(t => t.textContent || '').join('');
    const label = latex ? `[Equation: ${escapeHtml(latex)}]` : '[Equation]';
    const span = `<span class="equation">${label}</span>`;
    return span;
  }
  // Hyperlink → keep display text; resolve an external href when possible.
  if (tag === 'w:hyperlink') {
    const inner = renderXmlChildren(el, ctx, ast);
    if (!inner) return '';
    const anchor = el.getAttribute('w:anchor');
    const rId = el.getAttribute('r:id');
    let href = '';
    if (anchor) href = `#${anchor}`;
    else if (rId && ast) {
      const rel = (ast.relationships as Record<string, { target?: string }>)?.[rId];
      if (rel?.target) href = rel.target;
    }
    if (href) return `<a href="${escapeHtml(href)}">${inner}</a>`;
    return inner;
  }
  if (tag === 'w:br') return '<br/>';
  if (tag === 'w:t') return escapeHtml(el.textContent || '');
  if (tag === 'w:tab') return ' ';
  // Inline drawings are rendered by the node-level image branch — emitting
  // their (usually empty) text here would duplicate content.
  if (tag === 'w:drawing' || tag === 'wp:inline' || tag === 'wp:anchor') return '';
  // Unknown wrappers (smartTag, sdt, proofErr, …) → recurse so text survives.
  return renderXmlChildren(el, ctx, ast);
}

function equationLatexFromXml(xml: Element | null | undefined): string {
  if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return '';
  try {
    const mts = Array.from((xml as Element).getElementsByTagName('m:t'));
    if (mts.length > 0) return mts.map(t => t.textContent || '').join('');
  } catch { /* ignore */ }
  return '';
}

/** OX-07: dedupe header/footer zone texts (rId + filename keys alias the same part).
 * NOTE (D2 header over-dedupe review): text-keyed dedupe is intentional here,
 * not a bug. The parser stores the SAME part object under two keys
 * (matchedRelId + fileName sharing one reference), so alias collapse is
 * required; two genuinely distinct parts with byte-identical text are
 * visually identical in linear HTML export, so rendering the text twice
 * would be duplicate noise, not data loss. Dedupe-by-rId would render the
 * alias twice (worse). Kept as-is. */
function uniqueZoneTexts(map: Record<string, { text?: string }> | undefined): string[] {
  if (!map) return [];
  const out: string[] = [];
  const seen = new Set<string>();
  for (const entry of Object.values(map)) {
    const t = (entry?.text || '').trim();
    if (!t || seen.has(t)) continue;
    seen.add(t);
    out.push(t);
  }
  return out;
}

/** Multi-line code paragraphs render as <pre><code> with literal newlines
 * (mirrors the Markdown fenced-block gate in markdown.ts): a code named
 * style, or every text-bearing run monospace. Plain multi-line paragraphs
 * and single-line code keep their normal <p> rendering. */
function isCodeBlockParagraph(node: ExportNodeLike): boolean {
  const xml = node?.xml as Element | undefined;
  const metaStyle = String((metaRecordOf(node)['namedStyle'] as string | undefined) || (metaRecordOf(node)['style'] as string | undefined) || '').toLowerCase();
  let pStyleVal = '';
  if (xml && typeof (xml as unknown as Element).getElementsByTagName === 'function') {
    const pStyleEl = xml.getElementsByTagName('w:pStyle')[0];
    if (pStyleEl) pStyleVal = ((pStyleEl as Element).getAttribute('w:val') || '').toLowerCase();
  }
  const named = metaStyle || pStyleVal;
  if (named === 'code' || named === 'monospace' || named === 'inlinecode') return true;
  if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return false;
  const runs = Array.from(xml.getElementsByTagName('w:r'));
  let textRuns = 0;
  for (const r of runs) {
    // Accepted view: runs inside deletions carry no display text.
    let ancestor: Node | null = (r as unknown as Element).parentNode;
    let inDel = false;
    while (ancestor && ancestor !== xml) {
      if ((ancestor as Element).tagName === 'w:del') { inDel = true; break; }
      ancestor = ancestor.parentNode;
    }
    if (inDel) continue;
    const tEls = (r as Element).getElementsByTagName('w:t');
    let hasText = false;
    for (let i = 0; i < tEls.length; i++) {
      if (((tEls[i] as Element).textContent || '').length > 0) { hasText = true; break; }
    }
    if (!hasText) continue;
    textRuns++;
    if (!isMonospaceRun(r as Element)) return false;
  }
  return textRuns > 0;
}

/** Literal code content for <pre><code>: w:t text with w:br as newlines,
 * accepted view (insertions kept, deletions and field codes omitted). */
function codeBlockText(xml: Element | null | undefined): string {
  if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return '';
  let out = '';
  const walk = (el: Element): void => {
    for (let i = 0; i < el.childNodes.length; i++) {
      const ch = el.childNodes[i] as Element;
      if (!ch || (ch as unknown as Element).nodeType !== 1) continue;
      const tag = (ch as Element).tagName || '';
      if (tag === 'w:del' || tag === 'w:delText' || tag === 'w:fldChar' || tag === 'w:instrText') continue;
      if (tag === 'w:t') out += (ch as Element).textContent || '';
      else if (tag === 'w:br') out += '\n';
      else if (tag === 'w:tab') out += ' ';
      else if (tag === 'w:drawing' || tag === 'wp:inline' || tag === 'wp:anchor') continue;
      else walk(ch as Element);
    }
  };
  walk(xml);
  return out.replace(/[ \t]+\n/g, '\n').trim();
}

function renderParagraphContent(node: ExportNodeLike, ast?: DocumentAST): ParagraphRender {  const ctx = { footnoteIds: [] as string[], commentIds: [] as string[] };
  let html = '';
  if (node.xml && typeof (node.xml as unknown as Element).getElementsByTagName === 'function') {
    html = renderXmlChildren(node.xml as Element, ctx, ast);
  }
  if (!html) {
    html = escapeHtml(metaTextOf(node) || '');
  }
  if (ctx.commentIds.length > 0) {
    const label = ctx.commentIds.length === 1 && ctx.commentIds[0] !== '?'
      ? `[Comment ${escapeHtml(ctx.commentIds[0])}]`
      : `[Comments ${ctx.commentIds.map(escapeHtml).join(', ')}]`;
    html += ` <span class="comment-badge">${label}</span>`;
  }
  return { html, footnoteIds: ctx.footnoteIds, commentIds: ctx.commentIds };
}

/**
 * Callout nodes are single-cell tables in OOXML (see operations/callout.ts)
 * carrying `metadata.callout = { theme, title }`. They render as the
 * semantic `.callout` div the stylesheet already defines instead of a
 * 1x1 data table. Title comes from the marker; body is the marker text
 * minus the `Title: ` prefix, falling back to the cell paragraphs.
 */
export function renderCallout(node: ExportNodeLike): string {
  const metaRec = metaRecordOf(node);
  const raw = metaRec['callout'];
  const calloutRec: Record<string, unknown> =
    raw !== null && typeof raw === 'object' ? (raw as Record<string, unknown>) : {};
  const titleRaw = calloutRec['title'];
  const title = typeof titleRaw === 'string' ? titleRaw.trim() : '';
  let text = metaTextOf(node);
  if (title && text.startsWith(`${title}: `)) text = text.slice(title.length + 2).trim();
  if (!text) {
    const xml = node.xml as unknown as Element | undefined;
    if (xml && typeof (xml as Element).getElementsByTagName === 'function') {
      const paras = Array.from((xml as Element).getElementsByTagName('w:p'));
      const parts: string[] = [];
      for (const p of paras) {
        const runs = Array.from((p as Element).getElementsByTagName('w:t'));
        const joined = runs.map((t: Element) => t.textContent || '').join('').trim();
        if (joined) parts.push(joined);
      }
      text = parts.join(' ').trim();
      if (title && text.startsWith(`${title}: `)) text = text.slice(title.length + 2).trim();
      if (title && parts.length > 1 && parts[0] === title) text = parts.slice(1).join(' ').trim();
    }
  }
  if (!title && !text) return '';
  const titleHtml = title ? `<h4>${escapeHtml(title)}</h4>` : '';
  const bodyHtml = text ? `<p>${escapeHtml(text)}</p>` : '';
  return `<div class="callout">${titleHtml}${bodyHtml}</div>\n`;
}

/**
 * L-3: Render table node to HTML string, placing caption before or after table
 * based on node.metadata?.captionPosition ('above' | 'below').
 */
export function renderTable(node: ExportNodeLike): string {
  const metaRec = metaRecordOf(node);
  const metaHeaders = metaRec['headers'];
  const nodeHeaders = node.headers;
  let headers: string[] = Array.isArray(metaHeaders)
    ? (metaHeaders as unknown[]).map(String)
    : Array.isArray(nodeHeaders)
      ? (nodeHeaders as unknown[]).map(String)
      : [];
  const metaRows = metaRec['rows'];
  const nodeRows = node.rows;
  const coerceRows = (v: unknown): string[][] => {
    if (!Array.isArray(v)) return [];
    const arr = v as unknown[];
    if (arr.length === 0) return [];
    if (Array.isArray(arr[0])) return (arr as unknown[][]).map((r) => (r as unknown[]).map(String));
    return [];
  };
  let rows: string[][] = Array.isArray(metaRows)
    ? coerceRows(metaRows)
    : Array.isArray(nodeRows)
      ? coerceRows(nodeRows)
      : [];
  let cellColors: Array<Array<{ bg?: string; color?: string }>> = [];

  if (headers.length === 0 && rows.length === 0 && Array.isArray(nodeRows)) {
    const extractedRows: string[][] = [];
    for (const r of nodeRows as unknown[]) {
      if (Array.isArray(r)) {
        extractedRows.push((r as unknown[]).map(String));
      } else if (r && typeof r === 'object' && Array.isArray((r as { cells?: unknown }).cells)) {
        const cells = (r as { cells: unknown[] }).cells;
        extractedRows.push(cells.map((c: unknown) => {
          if (typeof c === 'string') return c;
          if (c !== null && typeof c === 'object') {
            const t = (c as Record<string, unknown>)['text'];
            return typeof t === 'string' ? t : '';
          }
          return '';
        }));
      }
    }
    if (extractedRows.length > 0) {
      headers = extractedRows[0];
      rows = extractedRows.slice(1);
    }
  }

  // Direct XML DOM table extraction with cell colors (<w:shd>, <w:color>)
  const nodeXmlEl = node.xml as unknown as Element | undefined;
  if (nodeXmlEl && typeof (nodeXmlEl as Element).getElementsByTagName === 'function') {
    const trs = Array.from((nodeXmlEl as Element).getElementsByTagName('w:tr'));
    const tableData: string[][] = [];
    const extractedColors: Array<Array<{ bg?: string; color?: string }>> = [];

    for (let rIdx = 0; rIdx < trs.length; rIdx++) {
      const tr = trs[rIdx] as Element;
      const tcs = Array.from(tr.getElementsByTagName('w:tc'));
      const rowCells: string[] = [];
      const rowColors: Array<{ bg?: string; color?: string }> = [];

      for (const tc of tcs) {
        const textEls = Array.from((tc as Element).getElementsByTagName('w:t'));
        const cellText = textEls.map((t: Element) => t.textContent || '').join(' ').trim();
        rowCells.push(cellText);

        // Shading (<w:shd>)
        let bg: string | undefined = undefined;
        const shdEl = (tc as Element).getElementsByTagName('w:shd')[0];
        if (shdEl) {
          const fill = shdEl.getAttribute('w:fill');
          if (fill && fill !== 'auto' && fill !== 'none' && fill !== 'clear') {
            bg = `#${fill.replace(/^#/, '')}`;
          }
        }

        // Text Color (<w:color>)
        let color: string | undefined = undefined;
        const colorEl = (tc as Element).getElementsByTagName('w:color')[0];
        if (colorEl) {
          const val = colorEl.getAttribute('w:val');
          if (val && val !== 'auto') {
            color = `#${val.replace(/^#/, '')}`;
          }
        } else if (bg && isColorDark(bg)) {
          color = '#FFFFFF';
        }

        rowColors.push({ bg, color });
      }

      if (rowCells.length > 0) {
        tableData.push(rowCells);
        extractedColors.push(rowColors);
      }
    }

    if (tableData.length > 0) {
      headers = tableData[0];
      rows = tableData.slice(1);
      cellColors = extractedColors;
    }
  }

  const captionRaw = node.caption || (metaRec['caption'] as string | undefined) || (metaRec['captionText'] as string | undefined);
  let caption: string | undefined = typeof captionRaw === 'string' ? captionRaw : undefined;
  if (!caption && nodeXmlEl) {
    const tblPr = (nodeXmlEl as Element).getElementsByTagName('w:tblPr')[0];
    if (tblPr) {
      const tblCaptionEl = tblPr.getElementsByTagName('w:tblCaption')[0];
      const tblTitleEl = tblPr.getElementsByTagName('w:tblTitle')[0];
      const tblDescEl = tblPr.getElementsByTagName('w:tblDescription')[0];
      caption = tblCaptionEl?.getAttribute('w:val') ||
                tblTitleEl?.getAttribute('w:val') ||
                tblDescEl?.getAttribute('w:val') ||
                undefined;
    }
  }

  const captionPositionRaw = node.captionPosition || (metaRec['captionPosition'] as string | undefined) || (metaRec['captionPlacement'] as string | undefined) || 'above';
  const captionPosition: string = typeof captionPositionRaw === 'string' ? captionPositionRaw : 'above';

  let tableHtml = '';
  if (headers.length > 0 || rows.length > 0) {
    const headerRow = headers.length > 0 ? headers : rows[0] || [];
    const dataRows = headers.length > 0 ? rows : rows.slice(1);
    const headerColors = cellColors[0] || [];
    const dataColors = cellColors.slice(1);

    tableHtml += `<table class="doc-table">\n<thead>\n<tr>`;
    headerRow.forEach((h: string, idx: number) => {
      const c = headerColors[idx];
      const bgStyle = c?.bg ? `background-color: ${c.bg};` : 'background-color: #F1F5F9;';
      const colorStyle = c?.color ? `color: ${c.color};` : 'color: #0F172A;';
      const styleAttr = ` style="${bgStyle} ${colorStyle}"`;
      tableHtml += `<th${styleAttr}>${escapeHtml(h)}</th>`;
    });
    tableHtml += `</tr>\n</thead>\n<tbody>\n`;
    dataRows.forEach((row: string[], rIdx: number) => {
      tableHtml += `<tr>`;
      const rowColorList = dataColors[rIdx] || [];
      row.forEach((c: string, cIdx: number) => {
        const colorInfo = rowColorList[cIdx];
        const bgStyle = colorInfo?.bg ? `background-color: ${colorInfo.bg};` : '';
        const colorStyle = colorInfo?.color ? `color: ${colorInfo.color};` : '';
        const styleAttr = (bgStyle || colorStyle) ? ` style="${bgStyle} ${colorStyle}"` : '';
        tableHtml += `<td${styleAttr}>${escapeHtml(c)}</td>`;
      });
      tableHtml += `</tr>\n`;
    });
    tableHtml += `</tbody>\n</table>\n`;
  }

  if (!tableHtml && !caption) return '';

  const captionHtml = caption ? `<p class="table-caption">${escapeHtml(caption)}</p>\n` : '';
  if (captionPosition === 'below') {
    return tableHtml + captionHtml;
  } else {
    return captionHtml + tableHtml;
  }
}

/**
 * R-3: footnote/endnote body ingestion — same sourcing and same `fn:`/`en:`
 * keying as the Markdown path's extractFootnoteBodies (ast.footnotes /
 * ast.endnotes first, then word/footnotes.xml + word/endnotes.xml via the
 * optional parser; separator ids and separator types skipped; footnoteRef /
 * endnoteRef runs excluded). Only in-text refs emit (orphan defs skipped
 * by construction — callers iterate footnoteOrder, never this map).
 */
function extractHtmlNoteBodies(parser?: DocxParser, ast?: DocumentAST): Map<string, string> {
  const bodies = new Map<string, string>();

  const ingestNoteMap = (src: unknown, prefix: 'fn' | 'en'): void => {
    if (!src) return;
    if (src instanceof Map) {
      for (const [k, v] of (src as Map<unknown, unknown>).entries()) {
        const rec = v as unknown as Record<string, unknown> | string;
        const rawText = typeof v === 'string' ? v : String((rec as Record<string, unknown>)?.['text'] ?? v ?? '');
        const text = rawText;
        if (text) bodies.set(`${prefix}:${String(k)}`, text.trim());
      }
    } else if (Array.isArray(src)) {
      for (const item of src) {
        const rec = item as unknown as Record<string, unknown>;
        if (item && rec['id'] != null) {
          const raw = rec['text'];
          bodies.set(`${prefix}:${String(rec['id'])}`, ((typeof raw === 'string' ? raw : '') as string).trim());
        }
      }
    } else if (typeof src === 'object') {
      for (const [k, v] of Object.entries(src as Record<string, unknown>)) {
        const rec = v as unknown as Record<string, unknown> | string;
        const rawText = typeof v === 'string' ? v : String((rec as Record<string, unknown>)?.['text'] ?? v ?? '');
        const text = rawText;
        if (text) bodies.set(`${prefix}:${String(k)}`, text.trim());
      }
    }
  };
  if (ast) {
    ingestNoteMap((ast as unknown as { footnotes?: unknown }).footnotes, 'fn');
    ingestNoteMap((ast as unknown as { endnotes?: unknown }).endnotes, 'en');
  }

  if (parser) {
    const parseNoteElements = (xmlContent: string, tagName: 'w:footnote' | 'w:endnote', prefix: 'fn' | 'en') => {
      assertNoXxe(xmlContent, tagName === 'w:footnote' ? 'word/footnotes.xml' : 'word/endnotes.xml');
      try {
        const doc = new DOMParser().parseFromString(xmlContent, 'text/xml');
        const elements = doc.getElementsByTagName(tagName);
        for (let i = 0; i < elements.length; i++) {
          const el = elements[i];
          const id = el.getAttribute('w:id');
          if (!id) continue;
          const numId = parseInt(id, 10);
          if (!isNaN(numId) && numId < 0) continue;
          const type = el.getAttribute('w:type');
          if (type === 'separator' || type === 'continuationSeparator') continue;

          const paras = Array.from(el.getElementsByTagName('w:p'));
          const pTexts: string[] = [];
          if (paras.length > 0) {
            for (const p of paras) {
              const runs = Array.from(p.getElementsByTagName('w:r'));
              const runParts: string[] = [];
              for (const r of runs) {
                if (r.getElementsByTagName('w:footnoteRef').length > 0 || r.getElementsByTagName('w:endnoteRef').length > 0) {
                  continue;
                }
                const ts = Array.from(r.getElementsByTagName('w:t'));
                for (const t of ts) {
                  runParts.push(t.textContent || '');
                }
                if (r.getElementsByTagName('w:br').length > 0) runParts.push('\n');
                if (r.getElementsByTagName('w:tab').length > 0) runParts.push(' ');
              }
              const pText = runParts.join('').trim();
              if (pText) pTexts.push(pText);
            }
          }
          let noteText = pTexts.join(' ').trim();
          if (!noteText) {
            const runs = Array.from(el.getElementsByTagName('w:r')).filter(
              r => r.getElementsByTagName('w:footnoteRef').length === 0 && r.getElementsByTagName('w:endnoteRef').length === 0
            );
            const ts: string[] = [];
            for (const r of runs) {
              for (const t of Array.from(r.getElementsByTagName('w:t'))) {
                ts.push(t.textContent || '');
              }
            }
            noteText = ts.join('').trim();
            if (!noteText) {
              noteText = Array.from(el.getElementsByTagName('w:t')).map(t => t.textContent || '').join('').trim();
            }
          }
          if (noteText) {
            bodies.set(`${prefix}:${id}`, noteText);
          }
        }
      } catch {
        /* best effort on xml parse */
      }
    };

    let footnotesXml: string | null = null;
    let endnotesXml: string | null = null;

    try {
      const caps = asFootnoteZipCapable(parser);
      if (typeof caps.getZipAdapter === 'function') {
        const adapter = caps.getZipAdapter();
        footnotesXml = adapter.readTextOrNull('word/footnotes.xml');
        endnotesXml = adapter.readTextOrNull('word/endnotes.xml');
      } else if (caps.zip && typeof caps.zip.getEntry === 'function') {
        const fnEntry = caps.zip.getEntry('word/footnotes.xml');
        if (fnEntry) footnotesXml = fnEntry.getData().toString('utf8');
        const enEntry = caps.zip.getEntry('word/endnotes.xml');
        if (enEntry) endnotesXml = enEntry.getData().toString('utf8');
      }
    } catch {
      /* ignore zip read error */
    }

    if (footnotesXml) parseNoteElements(footnotesXml, 'w:footnote', 'fn');
    if (endnotesXml) parseNoteElements(endnotesXml, 'w:endnote', 'en');
  }

  return bodies;
}

export function convertAstToHtml(
  ast: DocumentAST,
  title = 'Document',
  optionsOrParser?: { printMode?: boolean } | DocxParser,
  maybeParser?: DocxParser
): string {
  let options: { printMode?: boolean } | undefined;
  let parser: DocxParser | undefined;
  const looksLikeParser = (v: unknown): boolean => {
    if (!v || typeof v !== 'object') return false;
    const rec = v as Record<string, unknown>;
    return typeof rec['getZipAdapter'] === 'function' ||
      (rec['zip'] !== null && typeof rec['zip'] === 'object' &&
        typeof (rec['zip'] as Record<string, unknown>)['getEntry'] === 'function');
  };
  if (looksLikeParser(optionsOrParser)) {
    parser = optionsOrParser as DocxParser;
    if (looksLikeParser(maybeParser)) parser = maybeParser as DocxParser;
  } else {
    options = optionsOrParser as { printMode?: boolean } | undefined;
    if (looksLikeParser(maybeParser)) parser = maybeParser as DocxParser;
  }
  const noteBodies = extractHtmlNoteBodies(parser, ast);
  let bodyContent = '';
  const footnoteOrder: string[] = [];
  const noteFootnote = (ids: string[]): void => {
    for (const id of ids) {
      if (!footnoteOrder.includes(id)) footnoteOrder.push(id);
    }
  };
  // OX-07: consecutive list items group into one <ul>/<ol> (bare <li> is
  // invalid HTML and broke styling). A run flushes on style change
  // (ordered vs bulleted), on any non-list node, and at section end.
  let openListTag: string | null = null;
  const flushList = (): void => {
    if (openListTag) {
      bodyContent += openListTag === 'ol' ? `</ol>\n` : `</ul>\n`;
      openListTag = null;
    }
  };
  const listItemHtml = (node: ExportNodeLike): { tag: 'ol' | 'ul'; li: string; footnoteIds: string[] } => {
    const tag = isOrderedListItem(node as unknown as { metadata?: unknown; xml?: unknown }, ast) ? 'ol' : 'ul';
    const rendered = renderParagraphContent(node, ast);
    return { tag, li: `<li>${rendered.html || escapeHtml(metaTextOf(node) || '')}</li>\n`, footnoteIds: rendered.footnoteIds };
  };

  for (const section of ast.sections) {
    if (section.heading) {
      flushList();
      const level = Math.min(section.heading.level || 1, 6);
      const tag = `h${level}`;
      const align = getParagraphAlignment(section.heading.xml as Element);
      const alignStyle = align ? ` style="text-align: ${align};"` : '';
      const headingContent = renderParagraphContent(section.heading, ast);
      noteFootnote(headingContent.footnoteIds);
      bodyContent += `<${tag} class="heading"${alignStyle}>${headingContent.html}</${tag}>\n`;
    }

    for (const node of section.nodes) {
      if (node.type === 'paragraph' && !hasListNumbering(node)) {
        flushList();
        const text = node.metadata?.text || '';
        const pStyleEl = (node.xml as unknown as Element)?.getElementsByTagName ? (node.xml as Element).getElementsByTagName('w:pStyle')[0] : null;
        const pStyleVal = pStyleEl ? pStyleEl.getAttribute('w:val') || '' : '';
        const style = (node.metadata?.namedStyle || pStyleVal).toLowerCase();
        const align = getParagraphAlignment(node.xml as Element);
        const alignStyle = align ? ` style="text-align: ${align};"` : '';
        const rendered = renderParagraphContent(node, ast);
        noteFootnote(rendered.footnoteIds);
        const innerContent = rendered.html;
        if (isDivider(node, text)) {
          const dividerText = String((node.metadata as unknown as Record<string, unknown> | undefined)?.text || text || '').trim();
          if (dividerText && dividerText !== '---' && dividerText !== '***' && dividerText !== '___') {
            bodyContent += `<p class="paragraph">${escapeHtml(dividerText)}</p>\n`;
          }
          bodyContent += `<hr class="doc-divider"/>\n`;
        } else if (style.startsWith('heading') || style.startsWith('h1') || style.startsWith('h2') || style.startsWith('h3')) {
          const match = style.match(/\d+/);
          const level = match ? Math.min(parseInt(match[0], 10), 6) : 1;
          const tag = `h${level}`;
          bodyContent += `<${tag} class="heading"${alignStyle}>${innerContent}</${tag}>\n`;
        } else if (style === 'title' || style === 'tocheading') {
          bodyContent += `<h1 class="doc-title"${alignStyle}>${innerContent}</h1>\n`;
        } else if (isCodeBlockParagraph(node)) {
          // Multi-line code only: single-line code keeps its normal <p>.
          const raw = codeBlockText(node.xml as Element) || (node.metadata?.text || '');
          const trimmed = raw.trim();
          if (trimmed.includes('\n')) {
            const lang = codeBlockLanguage(node);
            const langAttr = /^[A-Za-z0-9_+#-]{1,24}$/.test(lang)
              ? ` class="language-${lang}"`
              : '';
            bodyContent += `<pre><code${langAttr}>${escapeHtml(trimmed)}</code></pre>\n`;
          } else if (text || innerContent) {
            bodyContent += `<p class="paragraph"${alignStyle}>${innerContent}</p>\n`;
          }
        } else if (text || innerContent) {
          bodyContent += `<p class="paragraph"${alignStyle}>${innerContent}</p>\n`;
        }
      } else if (node.type === 'table') {
        flushList();
        const calloutMeta = (node.metadata as unknown as Record<string, unknown> | undefined)?.['callout'];
        bodyContent += calloutMeta ? renderCallout(node) : renderTable(node);
      } else if ((node.type as string) === 'chart' || (node.type as string) === 'chart3d') {
        flushList();
        const meta = (node.metadata as unknown as Record<string, unknown> | undefined) || {};
        const strOr = (v: unknown, fallback: string): string => (typeof v === 'string' ? v : fallback);
        const arrOr = <T>(v: unknown, fallback: T): T => (Array.isArray(v) ? (v as T) : fallback);
        const chartData = (meta['chartData'] as { title?: string; categories?: string[]; series?: Array<{ name: string; values: number[]; color?: string }>; chartType?: string; holeSize?: number } | undefined) || {
          title: strOr(meta['title'], 'Data Analysis Chart'),
          categories: arrOr<string[]>(meta['categories'], ['Q1', 'Q2', 'Q3', 'Q4']),
          series: arrOr<Array<{ name: string; values: number[]; color?: string }>>(meta['series'], [{ name: 'Series 1', values: [25, 45, 60, 85], color: '2563EB' }]),
          chartType: strOr(meta['chartType'], 'bar3D')
        };
        bodyContent += renderVectorSvgChart(chartData);
      } else if (node.type === 'image' || (node.metadata as unknown as Record<string, unknown> | undefined)?.['hasImage']) {
        flushList();
        // OP-23: render embedded images as base64 <img>; fall back to alt placeholder.
        const img = resolveNodeImage(ast, node);
        if (img) {
          bodyContent += `<p class="doc-image"><img src="data:${img.mime};base64,${img.base64}" alt="${escapeHtml(img.alt)}" style="max-width:100%;height:auto;" /></p>\n`;
        } else {
          const altText = metaTextOf(node as unknown as ExportNodeLike).trim() || 'Image';
          bodyContent += `<p class="doc-image">[Image: ${escapeHtml(altText)}]</p>\n`;
        }
        // Paragraph-hosted drawings may also carry caption text alongside the image.
        const caption: string = metaTextOf(node as unknown as ExportNodeLike) || '';
        if (node.type !== 'image' && caption && caption !== '[Image]') {
          bodyContent += `<p class="paragraph">${escapeHtml(caption)}</p>\n`;
        }
      } else if (node.type === 'list' || (node.type === 'paragraph' && hasListNumbering(node))) {
        // OX-07: group consecutive items — never emit a bare <li>.
        const { tag, li, footnoteIds } = listItemHtml(node);
        noteFootnote(footnoteIds);
        if (openListTag !== tag) {
          flushList();
          bodyContent += tag === 'ol' ? `<ol class="doc-list">\n` : `<ul class="doc-list">\n`;
          openListTag = tag;
        }
        bodyContent += li;
      } else if ((node.type as string) === 'equation') {
        // E4: standalone equation node — native MathML display block.
        flushList();
        const latex: string = metaTextOf(node as unknown as ExportNodeLike) || equationLatexFromXml(node.xml as Element);
        bodyContent += `<p class="equation">${equationDisplayHtml(node.xml as Element, latex)}</p>\n`;
      } else if ((node.type as string) === 'field') {
        // OX-07: field node — display text only, codes omitted by the renderer.
        flushList();
        const rendered = renderParagraphContent(node, ast);
        noteFootnote(rendered.footnoteIds);
        const text: string = rendered.html || escapeHtml(metaTextOf(node as unknown as ExportNodeLike) || '[Field]');
        bodyContent += `<p class="paragraph field">${text}</p>\n`;
      } else if (isDivider(node)) {
        // Divider nodes render a horizontal rule; attached text stays above it.
        flushList();
        const dividerText = String(metaTextOf(node as unknown as ExportNodeLike) || '').trim();
        if (dividerText && dividerText !== '---' && dividerText !== '***' && dividerText !== '___') {
          bodyContent += `<p class="paragraph">${escapeHtml(dividerText)}</p>\n`;
        }
        bodyContent += `<hr class="doc-divider"/>\n`;
      } else {
        // OX-07 zero-silent-drop fallback: text when present, else an
        // explicit unsupported placeholder (loud, never vanished).
        flushList();
        const text = (node.metadata as unknown as Record<string, unknown> | undefined)?.['text'];
        if (typeof text === 'string' && text) {
          bodyContent += `<p class="paragraph">${escapeHtml(text)}</p>\n`;
        } else {
          bodyContent += `<p class="unsupported">[Unsupported node type: ${escapeHtml(String(node.type || 'unknown'))}]</p>\n`;
        }
      }
    }
    flushList();
  }

  // OX-07: headers/footers render deduped (empty zones omitted entirely).
  const headerTexts = uniqueZoneTexts(ast.headers);
  const footerTexts = uniqueZoneTexts(ast.footers);
  let zonesTop = '';
  for (const t of headerTexts) {
    zonesTop += `<header class="doc-header">${escapeHtml(t)}</header>\n`;
  }
  let zonesBottom = '';
  if (footnoteOrder.length > 0) {
    zonesBottom += `<section class="footnotes">\n<hr/>\n<ol>\n`;
    for (const key of footnoteOrder) {
      const sep = key.indexOf(':');
      const kind = sep >= 0 ? key.slice(0, sep) : 'fn';
      const id = sep >= 0 ? key.slice(sep + 1) : key;
      const label = kind === 'en' ? `en-${id}` : id;
      const liId = kind === 'en' ? `en-${escapeHtml(id)}` : `fn-${escapeHtml(id)}`;
      const body = noteBodies.get(key);
      const liText = body ? escapeHtml(body) : `[Footnote ${escapeHtml(label)}: see document]`;
      zonesBottom += `<li id="${liId}">${liText}</li>\n`;
    }
    zonesBottom += `</ol>\n</section>\n`;
  }
  for (const t of footerTexts) {
    zonesBottom += `<footer class="doc-footer">${escapeHtml(t)}</footer>\n`;
  }
  bodyContent = zonesTop + bodyContent + zonesBottom;

  const printStyles = options?.printMode ? `
    @page {
      margin: 20mm 15mm;
      size: auto;
    }
    body { margin: 0; padding: 0; max-width: 100%; color: #000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    h1, h2, h3, p, table, .doc-chart { page-break-inside: avoid; }
    .doc-table { border: 1px solid #333; }
    .doc-table th, .doc-table td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  ` : '';

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(title)}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; max-width: 850px; margin: 40px auto; padding: 0 20px; color: #1E293B; line-height: 1.6; }
    h1 { font-size: 24px; color: inherit; padding-bottom: 8px; margin-top: 32px; }
    h2 { font-size: 20px; color: inherit; margin-top: 24px; }
    h3 { font-size: 16px; color: inherit; }
    h4, h5, h6 { color: inherit; }
    p { margin: 10px 0; }
    .doc-table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px; }
    .doc-table th, .doc-table td { border: 1px solid #CBD5E1; padding: 8px 12px; text-align: left; }
    .doc-table th { background-color: #F1F5F9; color: #0F172A; font-weight: 600; }
    .doc-table tr:nth-child(even) td { background-color: #F8FAFC; }
    .table-caption { font-style: italic; color: #64748B; margin: 6px 0 12px 0; font-size: 12px; }
    .callout { border-left: 4px solid #4F46E5; background-color: #EEF2FF; padding: 14px 18px; border-radius: 4px; margin: 18px 0; }
    .callout h4 { margin: 0 0 6px 0; color: #312E81; }
    .callout p { margin: 0; color: #3730A3; }
    .doc-chart { margin: 24px auto; text-align: center; max-width: 640px; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .doc-list { margin: 10px 0 10px 24px; }
    .doc-header { color: #64748B; font-size: 12px; border-bottom: 1px solid #E2E8F0; padding-bottom: 8px; margin-bottom: 16px; }
    .doc-divider { border: none; border-top: 1px solid #CBD5E1; margin: 16px 0; }
    .doc-footer { color: #64748B; font-size: 12px; border-top: 1px solid #E2E8F0; padding-top: 8px; margin-top: 16px; }
    .footnotes { font-size: 13px; color: #475569; margin-top: 24px; }
    .footnote-ref { font-size: 11px; color: #4F46E5; }
    .comment-badge { font-size: 11px; color: #92400E; background: #FEF3C7; border-radius: 4px; padding: 1px 6px; }
    .equation { font-style: italic; color: #312E81; }
    ins { text-decoration: underline; text-decoration-color: #059669; }
    .tracked-del { text-decoration: line-through; color: #b30000; background-color: #ffe6e6; }
    .tracked-ins { text-decoration: underline; color: #006600; background-color: #e6ffe6; }
    .unsupported { color: #B45309; background: #FFFBEB; border: 1px dashed #F59E0B; padding: 8px 12px; }
    ${printStyles}
  </style>
</head>
<body>
${bodyContent}
</body>
</html>`;
}

function renderVectorSvgChart(chartData: { title?: string; categories?: string[]; series?: Array<{ name: string; values: number[]; color?: string; xValues?: number[]; unpairedX?: number; unpairedY?: number; chartType?: 'column' | 'line'; axis?: 'primary' | 'secondary' }>; chartType?: string; holeSize?: number }): string {
  // Never invent a heading for a part that carries no title; TOON and the AST
  // use the same provenance rule.
  const title = chartData.title || '';
  // OX-07: respect the stored chart type instead of always drawing bars.
  const kindRaw = String((chartData as { chartType?: unknown }).chartType || '').toLowerCase();
  const kind = kindRaw.includes('combo') ? 'combo'
    : kindRaw.includes('scatter') ? 'scatter'
      : kindRaw.includes('pie') || kindRaw.includes('doughnut') ? 'pie'
        : (kindRaw.includes('line') || kindRaw.includes('area')) ? 'line' : 'bar';
  const guardedSeries = chartData.series ?? [];
  const hasScatterX = (chartData.categories?.length ?? 0) > 0
    || guardedSeries.some((ser) => Array.isArray(ser.xValues) && ser.xValues.length > 0);
  // An x-only series (its y container is malformed/absent) still carries
  // reportable x points: render the scatter branch so it can say how many
  // points were omitted instead of replacing the chart with a blanket
  // unreadable-data notice.
  const hasXOnlySeries = guardedSeries.some((ser) =>
    Array.isArray(ser.xValues) && ser.xValues.length > 0 && (ser.unpairedX ?? 0) > 0);
  if (kind === 'scatter' && (guardedSeries.length === 0
      || (!hasXOnlySeries && guardedSeries.every((ser) => !Array.isArray(ser.values) || ser.values.length === 0))
      || !hasScatterX)) {
    const guardHeading = title
      ? `<div style="font-size: 13px; font-weight: 700; color: #1E3A8A; margin-bottom: 8px; font-family: -apple-system, sans-serif;">${escapeHtml(title)}</div>`
      : '';
    return `<div class="doc-chart">${guardHeading}
    <div class="unsupported">Scatter data not rendered: no readable x/y data in this chart part.</div></div>\n`;
  }
  const categories = chartData.categories && chartData.categories.length > 0 ? chartData.categories : ['Category 1', 'Category 2', 'Category 3'];
  const series = chartData.series && chartData.series.length > 0 ? chartData.series : [{ name: 'Series 1', values: [10, 20, 30], color: '2563EB' }];
  const defaultColors = ['#2563EB', '#10B981', '#F59E0B', '#8B5CF6', '#EC4899', '#06B6D4'];
  const combo = kind === 'combo';
  const barSeries = combo ? series.filter((s) => s.chartType !== 'line') : series;
  const comboLines = combo ? series.filter((s) => s.chartType === 'line') : [];
  const isStacked = kind === 'bar' && kindRaw.includes('stacked');
  const isPercentStacked = isStacked && kindRaw.includes('percentstacked');

  if (kind === 'pie') {
    return renderPieSvgChart(title, categories, series, defaultColors, kindRaw.includes('doughnut'), chartData.holeSize);
  }
  if (kind === 'scatter') {
    return renderScatterSvgChart(title, categories, series, defaultColors);
  }

  const primarySeries = series.filter((s) => s.axis !== 'secondary');
  const scaleSeries = combo && primarySeries.length > 0 ? primarySeries : series;
  const allVals = scaleSeries.flatMap(s => s.values || [0]);
  // Stacked bars scale to per-category sums, not to individual values.
  const stackPos = categories.map((_, cIdx) =>
    scaleSeries.reduce((sum, s) => sum + Math.max(0, Number(s.values[cIdx]) || 0), 0));
  const stackNeg = categories.map((_, cIdx) =>
    scaleSeries.reduce((sum, s) => sum + Math.min(0, Number(s.values[cIdx]) || 0), 0));
  const dataMax = isPercentStacked ? 1 : isStacked ? Math.max(...stackPos, 0) : Math.max(...allVals, 0);
  const dataMin = isPercentStacked ? 0 : isStacked ? Math.min(...stackNeg, 0) : Math.min(...allVals, 0);
  const range = (dataMax - dataMin) || 1;

  const svgWidth = 560;
  const svgHeight = 220;
  const paddingLeft = 50;
  const paddingRight = 30;
  const paddingTop = 40;
  const paddingBottom = 40;
  const plotWidth = svgWidth - paddingLeft - paddingRight;
  const plotHeight = svgHeight - paddingTop - paddingBottom;
  const yZero = paddingTop + (dataMax / range) * plotHeight;
  const yForVal = (val: number): number => yZero - (val / range) * plotHeight;

  const catCount = categories.length;
  const serCount = combo && barSeries.length > 0 ? barSeries.length : series.length;
  const groupWidth = plotWidth / catCount;
  const barWidth = isStacked
    ? Math.max(12, Math.min(36, groupWidth * 0.6))
    : Math.max(12, Math.min(36, (groupWidth * 0.7) / serCount));

  let barsSvg = '';
  let gridSvg = '';
  let catLabelsSvg = '';
  let legendSvg = '';

  // Y-Axis Grid Lines & Tick Labels (span the data range so value zero has a baseline)
  for (let i = 0; i <= 4; i++) {
    const stepped = dataMin + (i / 4) * (dataMax - dataMin);
    const y = yForVal(stepped);
    const val = isPercentStacked ? `${Math.round(stepped * 100)}%` : String(Math.round(stepped));
    gridSvg += `<line x1="${paddingLeft}" y1="${y}" x2="${svgWidth - paddingRight}" y2="${y}" stroke="#E2E8F0" stroke-width="1" stroke-dasharray="${i === 0 ? 'none' : '3,3'}"/>\n`;
    gridSvg += `<text x="${paddingLeft - 8}" y="${y + 3.5}" text-anchor="end" font-size="10" fill="#64748B" font-family="-apple-system, sans-serif">${val}</text>\n`;
  }
  if (dataMin < 0) {
    gridSvg += `<line x1="${paddingLeft}" y1="${yZero}" x2="${svgWidth - paddingRight}" y2="${yZero}" stroke="#94A3B8" stroke-width="1.5"/>\n`;
  }

  // Category Labels (shared by bar + line)
  categories.forEach((cat, cIdx) => {
    const groupCenterX = paddingLeft + (cIdx + 0.5) * groupWidth;
    catLabelsSvg += `<text x="${groupCenterX}" y="${svgHeight - paddingBottom + 18}" text-anchor="middle" font-size="11" fill="#475569" font-weight="500" font-family="-apple-system, sans-serif">${escapeHtml(cat)}</text>\n`;
  });

  if (kind === 'line') {
    const isArea = kindRaw.includes('area');
    const xFor = (cIdx: number): number => paddingLeft + (catCount === 1 ? plotWidth / 2 : (cIdx / (catCount - 1)) * plotWidth);
    const yFor = (val: number): number => yForVal(val);
    series.forEach((ser, sIdx) => {
      const rawColor = ser.color || defaultColors[sIdx % defaultColors.length];
      const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
      const pts = categories.map((_, cIdx) => `${xFor(cIdx)},${yFor(ser.values[cIdx] || 0)}`).join(' ');
      if (isArea) {
        const base = `${xFor(0)},${yZero} ${pts} ${xFor(catCount - 1)},${yZero}`;
        barsSvg += `<polygon points="${base}" fill="${color}" opacity="0.18" />\n`;
      }
      barsSvg += `<polyline points="${pts}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round">\n`;
      barsSvg += `<title>${escapeHtml(ser.name)}</title>\n</polyline>\n`;
      categories.forEach((_, cIdx) => {
        const val = ser.values[cIdx] || 0;
        barsSvg += `<circle cx="${xFor(cIdx)}" cy="${yFor(val)}" r="3.5" fill="${color}"><title>${escapeHtml(ser.name)}: ${val}</title></circle>\n`;
      });
    });
  } else if (isStacked) {
    categories.forEach((_cat, cIdx) => {
      const groupCenterX = paddingLeft + (cIdx + 0.5) * groupWidth;
      const total = Math.max(stackPos[cIdx], 1);
      let posBase = 0;
      let negBase = 0;
      series.forEach((ser, sIdx) => {
        const rawVal = Number(ser.values[cIdx]) || 0;
        const val = isPercentStacked ? Math.max(0, rawVal) / total : rawVal;
        const segH = (Math.abs(val) / range) * plotHeight;
        const base = val >= 0 ? posBase : negBase;
        const barX = groupCenterX - barWidth / 2;
        const barY = val >= 0
          ? yZero - ((Math.abs(base) + Math.abs(val)) / range) * plotHeight
          : yZero + (Math.abs(base) / range) * plotHeight;
        const rawColor = ser.color || defaultColors[sIdx % defaultColors.length];
        const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
        const label = isPercentStacked ? `${Math.round(val * 100)}%` : String(rawVal);
        barsSvg += `<rect x="${barX}" y="${barY}" width="${barWidth}" height="${segH}" rx="3" fill="${color}">
        <title>${escapeHtml(ser.name)}: ${rawVal}${isPercentStacked ? ` (${label})` : ''}</title>
      </rect>\n`;
        if (segH > 14) {
          barsSvg += `<text x="${barX + barWidth / 2}" y="${barY + segH / 2 + 3}" text-anchor="middle" font-size="9.5" fill="#FFFFFF" font-weight="600" font-family="-apple-system, sans-serif">${label}</text>\n`;
        }
        if (val >= 0) posBase += val; else negBase += val;
      });
    });
  } else {
  // Bars & Category Labels
  categories.forEach((cat, cIdx) => {
    const groupCenterX = paddingLeft + (cIdx + 0.5) * groupWidth;

    barSeries.forEach((ser, sIdx) => {
      const val = ser.values[cIdx] || 0;
      const barH = (Math.abs(val) / range) * plotHeight;
      const colorIndex = combo ? Math.max(0, series.indexOf(ser)) : sIdx;
      const rawColor = ser.color || defaultColors[colorIndex % defaultColors.length];
      const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
      const barOffset = (sIdx - (serCount - 1) / 2) * (barWidth + 4);
      const barX = groupCenterX + barOffset - barWidth / 2;
      const barY = val >= 0 ? yZero - barH : yZero;

      barsSvg += `<rect x="${barX}" y="${barY}" width="${barWidth}" height="${barH}" rx="3" fill="${color}">
        <title>${escapeHtml(ser.name)}: ${val}</title>
      </rect>\n`;
      if (barH > 16) {
        const labelY = val >= 0 ? barY - 4 : barY + barH + 12;
        barsSvg += `<text x="${barX + barWidth / 2}" y="${labelY}" text-anchor="middle" font-size="9.5" fill="#1E293B" font-weight="600" font-family="-apple-system, sans-serif">${val}</text>\n`;
      }
    });
  });
  } // end bar branch

  if (combo) {
    const comboX = (cIdx: number): number => paddingLeft + (cIdx + 0.5) * groupWidth;
    comboLines.forEach((ser) => {
      const ownScale = ser.axis === 'secondary';
      const vals = ser.values.map((v) => Number(v) || 0);
      const max = ownScale ? Math.max(...vals, 0) : dataMax;
      const min = ownScale ? Math.min(...vals, 0) : dataMin;
      const span = (max - min) || 1;
      const yFor = (val: number): number => paddingTop + ((max - val) / span) * plotHeight;
      const colorIndex = Math.max(0, series.indexOf(ser));
      const rawColor = ser.color || defaultColors[colorIndex % defaultColors.length];
      const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
      const points = categories.map((_, cIdx) => `${comboX(cIdx)},${yFor(vals[cIdx] || 0)}`).join(' ');
      barsSvg += `<polyline points="${points}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round">\n`;
      barsSvg += `<title>${escapeHtml(ser.name)}</title>\n</polyline>\n`;
      categories.forEach((_, cIdx) => {
        barsSvg += `<circle cx="${comboX(cIdx)}" cy="${yFor(vals[cIdx] || 0)}" r="3.5" fill="${color}"><title>${escapeHtml(ser.name)}: ${vals[cIdx] || 0}</title></circle>\n`;
      });
    });
  }

  // Legend
  let legendX = paddingLeft;
  series.forEach((ser, sIdx) => {
    const rawColor = ser.color || defaultColors[sIdx % defaultColors.length];
    const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
    legendSvg += `<rect x="${legendX}" y="12" width="10" height="10" rx="2" fill="${color}"/>\n`;
    legendSvg += `<text x="${legendX + 14}" y="20" font-size="10" fill="#475569" font-family="-apple-system, sans-serif">${escapeHtml(ser.name)}</text>\n`;
    legendX += ser.name.length * 7 + 35;
  });

  const heading = title
    ? `    <div style="font-size: 13px; font-weight: 700; color: #1E3A8A; margin-bottom: 8px; font-family: -apple-system, sans-serif;">${escapeHtml(title)}</div>\n`
    : '';
  return `<div class="doc-chart">
${heading}    <svg width="100%" height="220" viewBox="0 0 ${svgWidth} ${svgHeight}" style="overflow: visible;">
      ${gridSvg}
      ${barsSvg}
      ${catLabelsSvg}
      ${legendSvg}
    </svg>
  </div>\n`;
}

/** Scatter branch: numeric x per series against y values, both axes scaled
 * to the data range (no category bands, no clamping to a neighbouring x). */
function renderScatterSvgChart(
  title: string,
  categories: string[],
  series: Array<{ name: string; values: number[]; color?: string; xValues?: number[]; unpairedX?: number; unpairedY?: number }>,
  defaultColors: string[]
): string {
  const svgWidth = 560;
  const svgHeight = 220;
  const paddingLeft = 50;
  const paddingRight = 30;
  const paddingTop = 40;
  const paddingBottom = 40;
  const plotWidth = svgWidth - paddingLeft - paddingRight;
  const plotHeight = svgHeight - paddingTop - paddingBottom;

  const categoryX = categories.map((c) => Number(c));
  const sharedX = categoryX.length > 0 && categoryX.every((v) => Number.isFinite(v))
    ? categoryX
    : undefined;
  // A legal scatter can carry a different c:xVal list per series; read each
  // series' own list and never reuse the first list or the last value.
  // Never invent an x for a point the part does not carry: a malformed short
  // list leaves the point out and the renderer reports how many were omitted.
  const anyOwnX = series.some((ser) => Array.isArray(ser.xValues) && ser.xValues.length > 0);
  const seriesX: Array<Array<number | undefined>> = series.map((ser) => {
    const own = Array.isArray(ser.xValues) && ser.xValues.length > 0
      ? ser.xValues.map((value) => Number(value))
      : undefined;
    const usableOwn = own !== undefined && own.every((value) => Number.isFinite(value)) ? own : undefined;
    return ser.values.map((_, i) => {
      // A present-but-short or absent own list must not be topped up from the
      // shared list when other series carry their own x; only an all-shared
      // part uses the shared categories.
      const value = usableOwn !== undefined ? usableOwn[i] : anyOwnX ? undefined : sharedX?.[i];
      return value !== undefined && Number.isFinite(value) ? value : undefined;
    });
  });
  const xValues = seriesX.flat().filter((value): value is number => value !== undefined);
  const missingXPoints = seriesX.reduce(
    (count, xs, index) => count + xs.filter((value) => value === undefined).length
      + (typeof series[index]?.unpairedY === 'number' ? series[index].unpairedY! : 0),
    0
  );
  const missingYPoints = series.reduce(
    (count, ser) => count + (typeof ser.unpairedX === 'number' && ser.unpairedX > 0 ? ser.unpairedX : 0),
    0
  );
  const yValues = series.flatMap((s) => (s.values || []).map((v) => (Number.isFinite(Number(v)) ? Number(v) : 0)));
  const range = (values: number[]): [number, number] => {
    const finite = values.filter((value) => Number.isFinite(value));
    if (finite.length === 0) return [0, 1];
    const lo = Math.min(...finite);
    const hi = Math.max(...finite);
    const span = hi - lo;
    if (!Number.isFinite(span)) return [lo, hi];
    const pad = span === 0 ? (Math.abs(hi) || 1) * 0.1 : span * 0.1;
    let min = lo - pad;
    let max = hi + pad;
    if (!Number.isFinite(min) || !Number.isFinite(max)) {
      min = lo;
      max = hi;
    }
    if (min === max) {
      min -= 1;
      max += 1;
    }
    return [min, max];
  };
  const [xMin, xMax] = range(xValues);
  const [yMin, yMax] = range(yValues.length > 0 ? yValues : [0, 1]);
  const lerp = (min: number, max: number, fraction: number): number => {
    const span = max - min;
    if (Number.isFinite(span)) return min + fraction * span;
    const scale = Math.max(Math.abs(min), Math.abs(max), 1);
    return scale * ((min / scale) * (1 - fraction) + (max / scale) * fraction);
  };
  const fractionOf = (value: number, min: number, max: number): number => {
    const span = max - min;
    if (!Number.isFinite(span)) {
      const scale = Math.max(Math.abs(min), Math.abs(max), 1);
      const denominator = max / scale - min / scale;
      const scaled = denominator === 0 ? 0.5 : (value / scale - min / scale) / denominator;
      return Number.isFinite(scaled) ? Math.max(0, Math.min(1, scaled)) : 0.5;
    }
    if (span === 0) return 0.5;
    const f = (value - min) / span;
    return Number.isFinite(f) ? Math.max(0, Math.min(1, f)) : 0.5;
  };
  const xFor = (value: number): number => paddingLeft + fractionOf(value, xMin, xMax) * plotWidth;
  const yFor = (value: number): number => paddingTop + (1 - fractionOf(value, yMin, yMax)) * plotHeight;
  // Precision follows the axis span, not the magnitude: a 60-second span of
  // epoch timestamps must still render distinct labels and tooltip x values.
  const formatValue = (value: number, span: number): string => {
    if (!Number.isFinite(value)) return '0';
    if (value === 0) return '0';
    const abs = Math.abs(value);
    const step = Number.isFinite(span) && span > 0 ? span / 4 : abs || 1;
    if (abs >= 1e15 || abs < 1e-4) {
      // A double carries at most 17 significant decimal digits; use as many
      // as the span needs so 1e15 and 1e15+1 do not collapse to one label.
      const sig = Math.max(1, Math.min(17, Math.ceil(Math.log10(abs / step)) + 1));
      return value.toExponential(sig - 1).replace(/\.?0+e/, 'e');
    }
    const decimals = Math.max(0, Math.min(20, Math.ceil(-Math.log10(step))));
    const fixed = value.toFixed(decimals);
    return fixed.includes('.') ? fixed.replace(/0+$/, '').replace(/\.$/, '') : fixed;
  };
  const tick = (value: number, span: number): string => formatValue(value, span);
  const coord = (value: number, span: number): string => formatValue(value, span);

  let gridSvg = '';
  for (let i = 0; i <= 4; i++) {
    const xv = lerp(xMin, xMax, i / 4);
    const x = xFor(xv);
    gridSvg += `<line x1="${x}" y1="${paddingTop}" x2="${x}" y2="${svgHeight - paddingBottom}" stroke="#E2E8F0" stroke-width="1" stroke-dasharray="3,3"/>\n`;
    gridSvg += `<text x="${x}" y="${svgHeight - paddingBottom + 18}" text-anchor="middle" font-size="10" fill="#64748B" font-family="-apple-system, sans-serif">${tick(xv, xMax - xMin)}</text>\n`;
    const yv = lerp(yMin, yMax, i / 4);
    const y = yFor(yv);
    gridSvg += `<line x1="${paddingLeft}" y1="${y}" x2="${svgWidth - paddingRight}" y2="${y}" stroke="#E2E8F0" stroke-width="1" stroke-dasharray="${i === 0 ? 'none' : '3,3'}"/>\n`;
    gridSvg += `<text x="${paddingLeft - 8}" y="${y + 3.5}" text-anchor="end" font-size="10" fill="#64748B" font-family="-apple-system, sans-serif">${tick(yv, yMax - yMin)}</text>\n`;
  }

  let pointsSvg = '';
  series.forEach((ser, sIdx) => {
    const rawColor = ser.color || defaultColors[sIdx % defaultColors.length];
    const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
    const xs = seriesX[sIdx];
    const pts = ser.values.flatMap((v, i) => {
      const xv = xs[i];
      return xv === undefined ? [] : [`${xFor(xv)},${yFor(Number(v) || 0)}`];
    });
    if (pts.length > 0) {
      pointsSvg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${color}" stroke-width="2" opacity="0.65" stroke-linejoin="round" stroke-linecap="round">\n`;
      pointsSvg += `<title>${escapeHtml(ser.name)}</title>\n</polyline>\n`;
    }
    ser.values.forEach((v, i) => {
      const xv = xs[i];
      if (xv === undefined) return;
      const yv = Number(v) || 0;
      pointsSvg += `<circle cx="${xFor(xv)}" cy="${yFor(yv)}" r="3.5" fill="${color}"><title>${escapeHtml(ser.name)}: (${coord(xv, xMax - xMin)}, ${coord(yv, yMax - yMin)})</title></circle>\n`;
    });
  });
  if (missingXPoints > 0 || missingYPoints > 0) {
    const notes: string[] = [];
    if (missingXPoints > 0) {
      notes.push(!anyOwnX && sharedX === undefined && categories.length > 0
        ? 'x values are not numeric; scatter points not plotted'
        : `${missingXPoints} point${missingXPoints === 1 ? '' : 's'} omitted: no x value`);
    }
    if (missingYPoints > 0) {
      notes.push(`${missingYPoints} point${missingYPoints === 1 ? '' : 's'} omitted: no y value`);
    }
    pointsSvg += `<text x="${svgWidth - paddingRight}" y="${svgHeight - 6}" text-anchor="end" font-size="9" fill="#B45309" font-family="-apple-system, sans-serif">${notes.join('; ')}</text>\n`;
  }

  let legendSvg = '';
  let legendX = paddingLeft;
  series.forEach((ser, sIdx) => {
    const rawColor = ser.color || defaultColors[sIdx % defaultColors.length];
    const color = rawColor.startsWith('#') ? rawColor : `#${rawColor}`;
    legendSvg += `<rect x="${legendX}" y="12" width="10" height="10" rx="2" fill="${color}"/>\n`;
    legendSvg += `<text x="${legendX + 14}" y="20" font-size="10" fill="#475569" font-family="-apple-system, sans-serif">${escapeHtml(ser.name)}</text>\n`;
    legendX += ser.name.length * 7 + 35;
  });

  const heading = title
    ? `    <div style="font-size: 13px; font-weight: 700; color: #1E3A8A; margin-bottom: 8px; font-family: -apple-system, sans-serif;">${escapeHtml(title)}</div>\n`
    : '';
  return `<div class="doc-chart">
${heading}    <svg width="100%" height="220" viewBox="0 0 ${svgWidth} ${svgHeight}" style="overflow: visible;">
      ${gridSvg}
      ${pointsSvg}
      ${legendSvg}
    </svg>
  </div>\n`;
}

/** OX-07: pie/doughnut branch — first series over categories (shares must sum, not stack).
 * D2 pie-color fix: slices use the per-slice palette (legend matches slices
 * exactly via the shared `pieSliceColor` helper). A series-level `color` is
 * NOT smeared across all slices; it applies only to a single-slice pie (or
 * when explicit per-slice `colors`/`sliceColors` exist). Multi-series pies
 * render the first series plus an explicit note (documented placeholder) —
 * shares cannot stack, so series 1..n are listed, never silently dropped. */
function renderPieSvgChart(
  title: string,
  categories: string[],
  series: Array<{ name: string; values: number[]; color?: string; colors?: string[]; sliceColors?: string[] }>,
  defaultColors: string[],
  isDoughnut = false,
  holeSize?: number
): string {
  const first: { name: string; values: number[]; color?: string; colors?: string[]; sliceColors?: string[] } = series[0] || { name: 'Series 1', values: [] as number[] };
  const values = categories.map((_, i) => Math.max(0, Number(first.values[i]) || 0));
  const total = values.reduce((a, b) => a + b, 0);
  const cx = 170;
  const cy = 120;
  const r = 80;
  // c:holeSize is the hole diameter percentage of the chart diameter; the
  // default 50 therefore halves the outer radius. Clamp for corrupt input.
  const holePct = Math.max(1, Math.min(90, Math.round(holeSize ?? 50)));
  const innerR = isDoughnut ? (r * holePct) / 100 : 0;
  // Shared helper — slices AND legend call this so they can never diverge.
  const pieSliceColor = (i: number): string => {
    const explicitSlice =
      (Array.isArray(first.colors) && first.colors[i]) ||
      (Array.isArray(first.sliceColors) && first.sliceColors[i]);
    if (explicitSlice) {
      const c = String(explicitSlice);
      return c.startsWith('#') ? c : `#${c}`;
    }
    // Single-slice pie may honor the series color; multi-slice always uses
    // the palette so every slice is distinguishable.
    if (values.length === 1 && first.color) {
      const c = String(first.color);
      return c.startsWith('#') ? c : `#${c}`;
    }
    const raw = defaultColors[i % defaultColors.length];
    return String(raw).startsWith('#') ? String(raw) : `#${raw}`;
  };
  let slices = '';
  let labels = '';
  let legend = '';
  if (total <= 0) {
    slices = innerR > 0
      ? `<path fill-rule="evenodd" d="M ${cx - r} ${cy} a ${r} ${r} 0 1 0 ${r * 2} 0 a ${r} ${r} 0 1 0 ${-r * 2} 0 Z M ${cx - innerR} ${cy} a ${innerR} ${innerR} 0 1 0 ${innerR * 2} 0 a ${innerR} ${innerR} 0 1 0 ${-innerR * 2} 0 Z" fill="#E2E8F0" />`
      : `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#E2E8F0" />`;
    labels = `<text x="${cx}" y="${cy}" text-anchor="middle" font-size="11" fill="#64748B">no data</text>`;
  } else {
    let angle = -Math.PI / 2;
    values.forEach((val, i) => {
      const frac = val / total;
      if (frac <= 0) return;
      const a0 = angle;
      const a1 = angle + frac * Math.PI * 2;
      angle = a1;
      const color = pieSliceColor(i);
      const large = frac > 0.5 ? 1 : 0;
      const x0 = cx + r * Math.cos(a0);
      const y0 = cy + r * Math.sin(a0);
      const x1 = cx + r * Math.cos(a1);
      const y1 = cy + r * Math.sin(a1);
      // Doughnut slices are annular sectors (outer arc + inner arc), so the
      // hole is real geometry, not a pie with a background circle on top.
      const piePath = `M ${cx} ${cy} L ${x0.toFixed(1)} ${y0.toFixed(1)} A ${r} ${r} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} Z`;
      const fullRing = innerR > 0 && frac >= 0.999999;
      const donutPath = innerR === 0
        ? piePath
        : fullRing
          ? `M ${cx - r} ${cy} a ${r} ${r} 0 1 0 ${r * 2} 0 a ${r} ${r} 0 1 0 ${-r * 2} 0 Z M ${cx - innerR} ${cy} a ${innerR} ${innerR} 0 1 0 ${innerR * 2} 0 a ${innerR} ${innerR} 0 1 0 ${-innerR * 2} 0 Z`
          : `M ${x0.toFixed(1)} ${y0.toFixed(1)} A ${r} ${r} 0 ${large} 1 ${x1.toFixed(1)} ${y1.toFixed(1)} L ${(cx + innerR * Math.cos(a1)).toFixed(1)} ${(cy + innerR * Math.sin(a1)).toFixed(1)} A ${innerR} ${innerR} 0 ${large} 0 ${(cx + innerR * Math.cos(a0)).toFixed(1)} ${(cy + innerR * Math.sin(a0)).toFixed(1)} Z`;
      slices += `<path d="${donutPath}"${fullRing ? ' fill-rule="evenodd"' : ''} fill="${color}" stroke="#FFFFFF" stroke-width="2"><title>${escapeHtml(categories[i])}: ${val}</title></path>\n`;
      if (frac >= 0.08) {
        const mid = (a0 + a1) / 2;
        const labelR = innerR > 0 ? (r + innerR) / 2 : r * 0.68;
        const lx = cx + labelR * Math.cos(mid);
        const ly = cy + labelR * Math.sin(mid);
        labels += `<text x="${lx.toFixed(1)}" y="${ly.toFixed(1)}" text-anchor="middle" font-size="10" fill="#FFFFFF" font-weight="700">${Math.round(frac * 100)}%</text>\n`;
      }
    });
  }
  const legendX = 290;
  categories.forEach((cat, i) => {
    const color = pieSliceColor(i);
    const y = 60 + i * 22;
    legend += `<rect x="${legendX}" y="${y}" width="10" height="10" rx="2" fill="${color}"/>\n`;
    legend += `<text x="${legendX + 14}" y="${y + 8}" font-size="10" fill="#475569" font-family="-apple-system, sans-serif">${escapeHtml(cat)} (${values[i] ?? 0})</text>\n`;
  });
  void legendX;
  const shapeName = isDoughnut ? 'Doughnut' : 'Pie';
  const multiNote =
    series.length > 1
      ? `<div style="font-size: 11px; color: #64748B; margin-top: 6px; font-family: -apple-system, sans-serif;">${shapeName} shows “${escapeHtml(first.name || 'Series 1')}” only (${series.length} series provided: ${series.map((s: { name: string }) => escapeHtml(s.name || 'unnamed')).join(', ')}). See Markdown export for the full multi-series table.</div>`
      : '';
  const pieHeading = title
    ? `    <div style="font-size: 13px; font-weight: 700; color: #1E3A8A; margin-bottom: 8px; font-family: -apple-system, sans-serif;">${escapeHtml(title)}</div>\n`
    : '';
  return `<div class="doc-chart">
${pieHeading}    <svg width="100%" height="220" viewBox="0 0 560 220" style="overflow: visible;">
      ${slices}
      ${labels}
      ${legend}
    </svg>
    ${multiNote}
  </div>\n`;
}

function isColorDark(hex: string): boolean {
  const cleanHex = hex.replace(/^#/, '');
  if (cleanHex.length !== 6) return false;
  const r = parseInt(cleanHex.substring(0, 2), 16);
  const g = parseInt(cleanHex.substring(2, 4), 16);
  const b = parseInt(cleanHex.substring(4, 6), 16);
  const luma = 0.299 * r + 0.587 * g + 0.114 * b;
  return luma < 140;
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
