import { DocumentAST, Node, NodeMetadata } from '../ast';
import { DocxParser } from '../parser';
import { DOMParser } from '@xmldom/xmldom';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { ommlToLatex } from '../ooxml/omml';

/** Loose node view for exporters (mirrors html.ts ExportNodeLike; tolerates
 * partial/test-built nodes such as `{ xml }` passed for callout bodies). */
export type MarkdownNodeLike = {
  id?: string;
  type?: string;
  caption?: unknown;
  captionPosition?: unknown;
  headers?: unknown;
  rows?: unknown;
  xml?: unknown;
  metadata?: unknown;
};

function metaRecordOf(node: MarkdownNodeLike): Record<string, unknown> {
  if (typeof node.metadata === 'object' && node.metadata !== null) {
    return node.metadata as unknown as Record<string, unknown>;
  }
  return {};
}

function metaTextOf(node: MarkdownNodeLike): string {
  const t: unknown = metaRecordOf(node)['text'];
  if (typeof t === 'string') return t;
  if (t) return String(t);
  return '';
}

/** Narrow the optional parser zip fallback without `any` (mirrors html.ts
 * asFootnoteZipCapable; DocxParser.getZipAdapter covers the primary path). */
interface ZipEntryReader {
  getEntry: (p: string) => { getData: () => { toString: (enc: string) => string } } | null | undefined;
}

function parserZipOf(parser: unknown): ZipEntryReader | null {
  if (typeof parser !== 'object' || parser === null) return null;
  const zip = (parser as { zip?: unknown }).zip;
  if (typeof zip !== 'object' || zip === null) return null;
  const getEntry = (zip as { getEntry?: unknown }).getEntry;
  if (typeof getEntry !== 'function') return null;
  return { getEntry: (getEntry as ZipEntryReader['getEntry']).bind(zip) };
}

/** OP-23: resolve embedded image bytes for a node via a:blip r:embed -> relationships -> mediaFiles. */
function resolveNodeImage(ast: DocumentAST, node: MarkdownNodeLike): { base64: string; mime: string; alt: string; data: Buffer } | null {
  try {
    const xml = node?.xml as Element | undefined;
    let embedId: string | null = null;
    // Alt text is the drawing description; when a drawing carries none, the
    // fallback is the literal 'Image' — never the opaque session node id.
    let alt = metaTextOf(node).trim() || 'Image';
    if (alt === '[Image]') alt = 'Image';
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
    if (!buf) return target ? { base64: '', mime: '', alt, data: Buffer.alloc(0) } : null;
    const data = Buffer.isBuffer(buf) ? buf : Buffer.from(buf as Uint8Array);
    const ext = (base.split('.').pop() || 'png').toLowerCase();
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
      : ext === 'gif' ? 'image/gif'
      : ext === 'webp' ? 'image/webp'
      : ext === 'bmp' ? 'image/bmp'
      : ext === 'svg' ? 'image/svg+xml'
      : 'image/png';
    return { base64: data.toString('base64'), mime, alt, data };
  } catch {
    return null;
  }
}

/**
 * OX-07 exporter drop policy (mirrors html.ts; Markdown shapes differ).
 *
 * - footnotes/endnotes → inline `[^id]` (footnotes) / `[^en-id]` (endnotes)
 *   + matching defs with real bodies from word/footnotes.xml and
 *   word/endnotes.xml. Keys are namespaced (`fn:` vs `en:`) so a shared
 *   numeric id never collides; defs emit only for in-text refs (orphan
 *   definitions are skipped, never appended).
 * - comments → `> [Comment …]` quote line after the covered paragraph; the
 *   covered text itself is always preserved.
 * - fields + TOC → display text only; `w:instrText`/`w:fldChar` codes omitted.
 * - tracked changes → accepted view: insertions kept, deletions omitted.
 * - headers/footers → `*Header: …*` / `*Footer: …*` lines, deduped.
 * - equations → DISPLAY equations render as `$$` blocks; INLINE math keeps
 *   `$...$` (fallback `[Equation]`).
 * - images without bytes → `[Image: alt]` fallback text (never bare `![alt]`).
 * - unknown node types → `metadata.text` when present, else an explicit
 *   `[Unsupported node type: X]` line. Nothing drops silently.
 */
interface MdParagraph {
  text: string;
  footnoteIds: string[];
  commentIds: string[];
}

function mdChildrenText(
  parent: Element,
  ctx: { footnoteIds: string[]; commentIds: string[] },
  ast?: DocumentAST,
  raw = false
): string {
  let out = '';
  for (let i = 0; i < parent.childNodes.length; i++) {
    const ch = parent.childNodes[i] as Element;
    if (!ch || ch.nodeType !== 1) continue;
    out += mdChildText(ch, ctx, ast, raw);
  }
  return out;
}

export function isMonospaceRun(el: Element): boolean {
  if (!el || typeof el.getElementsByTagName !== 'function') return false;
  const rPr = el.getElementsByTagName('w:rPr')[0];
  if (!rPr) return false;

  const rStyle = rPr.getElementsByTagName('w:rStyle')[0];
  if (rStyle) {
    const s = (rStyle.getAttribute('w:val') || '').toLowerCase();
    if (s.includes('code') || s.includes('mono') || s.includes('verbatim') || s.includes('source') || s.includes('macro')) {
      return true;
    }
  }

  const rFonts = rPr.getElementsByTagName('w:rFonts')[0];
  if (rFonts) {
    const fonts = [
      rFonts.getAttribute('w:ascii'),
      rFonts.getAttribute('w:hAnsi'),
      rFonts.getAttribute('w:cs'),
      rFonts.getAttribute('w:eastAsia'),
      rFonts.getAttribute('w:rFonts'),
    ].filter(Boolean) as string[];

    const monoPattern = /consolas|courier|mono|code|terminal|fixedsys|menlo|fira|source\s*code/i;
    for (const font of fonts) {
      if (monoPattern.test(font)) {
        return true;
      }
    }
  }

  return false;
}

function isBoldRun(el: Element): boolean {
  if (!el || typeof el.getElementsByTagName !== 'function') return false;
  const rPr = el.getElementsByTagName('w:rPr')[0];
  if (!rPr) return false;
  const b = rPr.getElementsByTagName('w:b')[0];
  if (!b) return false;
  const val = (b.getAttribute('w:val') || '').toLowerCase();
  return !(val === '0' || val === 'false' || val === 'off' || val === 'none');
}

function isItalicRun(el: Element): boolean {
  if (!el || typeof el.getElementsByTagName !== 'function') return false;
  const rPr = el.getElementsByTagName('w:rPr')[0];
  if (!rPr) return false;
  const i = rPr.getElementsByTagName('w:i')[0];
  if (!i) return false;
  const val = (i.getAttribute('w:val') || '').toLowerCase();
  return !(val === '0' || val === 'false' || val === 'off' || val === 'none');
}

function mdEscapeRunText(s: string): string {
  return s
    .replace(/\\/g, '\\\\')
    .replace(/\*/g, '\\*')
    .replace(/_/g, '\\_')
    .replace(/`/g, '\\`')
    .replace(/\[/g, '\\[')
    .replace(/\]/g, '\\]');
}

function mdStyledSpan(text: string, bold: boolean, italic: boolean, code: boolean): string {
  const normalized = text.replace(/\r\n|\r|\n/g, ' ');
  if (!normalized) return normalized;
  if (code) {
    const match = normalized.match(/^(\s*)([\s\S]*?)(\s*)$/);
    if (!match || match[2].length === 0) return normalized;
    const inner = mdCodeSpan(match[2]);
    const withEm = italic ? `*${inner}*` : inner;
    const withStrong = bold ? `**${withEm}**` : withEm;
    return `${match[1]}${withStrong}${match[3]}`;
  }
  const escaped = mdEscapeRunText(normalized);
  const match = escaped.match(/^(\s*)([\s\S]*?)(\s*)$/);
  if (!match || match[2].length === 0) return normalized;
  const core = match[2];
  if (!bold && !italic) return `${match[1]}${core}${match[3]}`;
  const inner = italic ? `*${core}*` : core;
  const outer = bold ? `**${inner}**` : inner;
  return `${match[1]}${outer}${match[3]}`;
}

function mdChildText(
  el: Element,
  ctx: { footnoteIds: string[]; commentIds: string[] },
  ast?: DocumentAST,
  raw = false
): string {
  const tag = el.tagName || '';
  if (tag === 'w:del' || tag === 'w:delText') return '';
  if (tag === 'w:ins') return mdChildrenText(el, ctx, ast, raw);
  if (tag === 'w:r') {
    // A run holding only field codes (w:instrText, no w:t) yields ''.
    // Runs can also host inline specials — surface those (never dropped).
    const parts: string[] = [];
    const isMono = isMonospaceRun(el);
    const isBold = isBoldRun(el);
    const isItalic = isItalicRun(el);
    let runText = '';

    const flushText = () => {
      if (!runText) return;
      if (raw) {
        parts.push(runText);
      } else if (isMono || isBold || isItalic) {
        if (runText.trim().length > 0) {
          parts.push(mdStyledSpan(runText, isBold, isItalic, isMono));
        } else {
          parts.push(runText.replace(/\r\n|\r|\n/g, ' '));
        }
      } else {
        const normalized = runText.replace(/\r\n|\r|\n/g, ' ');
        parts.push(mdEscapeRunText(normalized));
      }
      runText = '';
    };

    for (let i = 0; i < el.childNodes.length; i++) {
      const c = el.childNodes[i] as Element;
      if (!c || c.nodeType !== 1) continue;
      if ((c as Element).tagName === 'w:t') {
        runText += (c as Element).textContent || '';
      } else if ((c as Element).tagName === 'w:tab') {
        runText += ' ';
      } else if ((c as Element).tagName === 'w:br') {
        flushText();
        parts.push('\n');
      } else {
        flushText();
        if ((c as Element).tagName === 'w:footnoteReference') {
          const id = (c as Element).getAttribute('w:id') || '?';
          const key = `fn:${id}`;
          if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
          parts.push(`[^${id}]`);
        } else if ((c as Element).tagName === 'w:endnoteReference') {
          const id = (c as Element).getAttribute('w:id') || '?';
          const key = `en:${id}`;
          if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
          parts.push(`[^en-${id}]`);
        } else if ((c as Element).tagName === 'w:commentReference') {
          const id = (c as Element).getAttribute('w:id') || '?';
          if (!ctx.commentIds.includes(id)) ctx.commentIds.push(id);
        } else if ((c as Element).tagName === 'm:oMath' || (c as Element).tagName === 'm:oMathPara') {
          const latex = mdMathLatex(c as Element);
          if (!latex) parts.push('[Equation]');
          else if ((c as Element).tagName === 'm:oMathPara') parts.push(`\n$$\n${latex}\n$$\n`);
          else parts.push(`\\(${latex}\\)`);
        }
      }
    }
    flushText();
    return parts.join('');
  }
  if (tag === 'w:footnoteReference') {
    const id = el.getAttribute('w:id') || '?';
    const key = `fn:${id}`;
    if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
    return `[^${id}]`;
  }
  if (tag === 'w:endnoteReference') {
    const id = el.getAttribute('w:id') || '?';
    const key = `en:${id}`;
    if (!ctx.footnoteIds.includes(key)) ctx.footnoteIds.push(key);
    return `[^en-${id}]`;
  }
  if (tag === 'w:commentRangeStart' || tag === 'w:commentRangeEnd' || tag === 'w:commentReference') {
    const id = el.getAttribute('w:id') || '';
    if (id && !ctx.commentIds.includes(id)) ctx.commentIds.push(id);
    else if (!id && !ctx.commentIds.includes('?')) ctx.commentIds.push('?');
    return '';
  }
  if (tag === 'w:fldChar' || tag === 'w:instrText') return '';
  if (tag === 'w:fldSimple') return mdChildrenText(el, ctx, ast, raw);
  if (tag === 'm:oMathPara') {
    const latex = mdMathLatex(el);
    return latex ? `\n$$\n${latex}\n$$\n` : '[Equation]';
  }
  if (tag === 'm:oMath') {
    const latex = mdMathLatex(el);
    return latex ? `\\(${latex}\\)` : '[Equation]';
  }
  if (tag === 'w:hyperlink') {
    const inner = mdChildrenText(el, ctx, ast, raw);
    if (!inner) return '';
    const anchor = el.getAttribute('w:anchor');
    const rId = el.getAttribute('r:id');
    let href = '';
    if (anchor) href = `#${anchor}`;
    else if (rId && ast) {
      const rel = (ast.relationships as Record<string, { target?: string }>)?.[rId];
      if (rel?.target) href = rel.target;
    }
    // Unresolvable (missing/empty rel, no anchor) degrades to plain text.
    if (!href) return inner;
    return /[\s()]/.test(href) ? `[${inner}](<${href}>)` : `[${inner}](${href})`;
  }
  if (tag === 'w:br') return '\n';
  if (tag === 'w:t') {
    const content = el.textContent || '';
    if (raw) return content;
    return mdEscapeRunText(content.replace(/\r\n|\r|\n/g, ' '));
  }
  if (tag === 'w:tab') return ' ';
  if (tag === 'w:drawing' || tag === 'wp:inline' || tag === 'wp:anchor') return '';
  return mdChildrenText(el, ctx, ast, raw);
}

/**
 * LaTeX for one formula: an m:oMath, or an element that holds exactly one. Several formulas in one element, or a
 * formula the converter does not emit, keep the flat m:t text (the former behavior), so nothing is dropped.
 */
function mdMathLatex(el: Element): string {
  try {
    const maths = el.tagName === 'm:oMath' ? [el] : Array.from(el.getElementsByTagName('m:oMath')) as Element[];
    if (maths.length === 1) {
      const latex = ommlToLatex(maths[0]);
      if (latex !== null) return latex;
    }
  } catch { /* fall back to the flat text */ }
  return Array.from(el.getElementsByTagName('m:t')).map(t => (t as Element).textContent || '').join('');
}

function mdEquationLatex(xml: Element | null | undefined): string {
  if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return '';
  try {
    if ((xml as Element).getElementsByTagName('m:t').length > 0) return mdMathLatex(xml as Element);
  } catch { /* ignore */ }
  return '';
}

function mdParagraph(node: MarkdownNodeLike, ast?: DocumentAST): MdParagraph {
  const ctx = { footnoteIds: [] as string[], commentIds: [] as string[] };
  let text = '';
  const nodeXml = node.xml as Element | undefined;
  if (nodeXml && typeof (nodeXml as unknown as Element).getElementsByTagName === 'function') {
    // Accepted view: XML-derived display text skips deletions + field codes.
    text = mdChildrenText(nodeXml as Element, ctx, ast).replace(/[ \t]+\n/g, '\n').trim();
  }
  if (!text) text = (metaTextOf(node) || '').trim();
  return { text, footnoteIds: ctx.footnoteIds, commentIds: ctx.commentIds };
}

/** True when a paragraph carries list numbering (w:numPr). The parser keeps
 * such paragraphs as type 'paragraph', so exporters must check the XML —
 * testing node.type alone misses every real-world list. Shared with html.ts. */
export function hasListNumbering(node: MarkdownNodeLike): boolean {
  try {
    const xml = node?.xml as Element | undefined;
    if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return false;
    return (xml as Element).getElementsByTagName('w:numPr').length > 0;
  } catch {
    return false;
  }
}

/** Bullet vs ordered for a list node or a numbered paragraph. Explicit
 * metadata wins; otherwise the numId resolves through ast.numbering to the
 * level's w:numFmt (anything but 'bullet' counts as ordered). Unresolvable
 * numbering defaults to bullet. Shared with html.ts. */
export function isOrderedListItem(node: MarkdownNodeLike, ast?: DocumentAST): boolean {
  const style = String(metaRecordOf(node)?.['listStyle'] || '').toLowerCase();
  if (style === 'decimal' || style === 'numbered' || style === 'ordered') return true;
  if (style === 'bullet') return false;
  try {
    const xml = node?.xml as Element | undefined;
    const numPr = xml?.getElementsByTagName?.('w:numPr')?.[0] as Element | undefined;
    const numId = numPr?.getElementsByTagName?.('w:numId')?.[0]?.getAttribute?.('w:val');
    if (numId == null || numId === '') return false;
    const concrete = ast?.numbering?.concreteNum as Record<string, Element> | undefined;
    const abstract = ast?.numbering?.abstractNum as Record<string, Element> | undefined;
    const numEl = concrete?.[numId];
    const absId = numEl?.getElementsByTagName?.('w:abstractNumId')?.[0]?.getAttribute?.('w:val');
    const absEl = (absId != null && absId !== '') ? abstract?.[absId] : undefined;
    if (!absEl) return false;
    const ilvl = numPr?.getElementsByTagName?.('w:ilvl')?.[0]?.getAttribute?.('w:val');
    const lvls = Array.from(absEl.getElementsByTagName?.('w:lvl') || []) as Element[];
    let lvlEl: Element | undefined;
    if (ilvl != null && ilvl !== '') {
      lvlEl = lvls.find(l => l.getAttribute?.('w:ilvl') === ilvl);
    }
    const fmt = lvlEl?.getElementsByTagName?.('w:numFmt')?.[0]?.getAttribute?.('w:val')
      || absEl.getElementsByTagName?.('w:numFmt')?.[0]?.getAttribute?.('w:val');
    if (!fmt) return false;
    return fmt !== 'bullet';
  } catch {
    return false;
  }
}

/** Dedupe header/footer zone texts (rId + filename keys alias one part).
 * NOTE (D2 header over-dedupe review): text-keyed dedupe intentional — same
 * rationale as html.ts uniqueZoneTexts. Identical texts are visually
 * identical in linear Markdown; rendering twice is noise, not loss. */
function mdZoneTexts(map: Record<string, { text?: string }> | undefined): string[] {
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

function mdEscapePipes(s: string): string {
  return s.replace(/\|/g, '\\|');
}

/** Wrap literal text as a GFM code span, escaping pipes and neutralizing
 * internal backticks by lengthening the fence (`` `x` `` style). */
function mdCodeSpan(s: string): string {
  const escaped = s.replace(/\|/g, '\\|');
  if (!escaped.includes('`')) return `\`${escaped}\``;
  const runs = escaped.match(/`+/g) || [];
  const longest = Math.max(...runs.map(r => r.length));
  const fence = '`'.repeat(longest + 1);
  return `${fence} ${escaped} ${fence}`;
}

/** Wrap multi-line code content as a fenced block. Pipes stay literal
 * inside fences (no GFM-table escape); the fence lengthens past any
 * backtick run in the content so the block never terminates early. A fence
 * language (restored from the code bookmark the insert path writes) rides
 * on the opening fence; blocks without one export bare. */
function mdFencedBlock(content: string, lang = ''): string {
  const runs = content.match(/`+/g) || [];
  const longest = runs.reduce((m, r) => Math.max(m, r.length), 0);
  const fence = '`'.repeat(Math.max(3, longest + 1));
  const tag = /^[A-Za-z0-9_+#-]{1,24}$/.test(lang) ? lang : '';
  return `${fence}${tag}\n${content}\n${fence}`;
}

/** Fence language for a code-block paragraph, read from the bookmark the
 * insert path writes (`_AstraCode_<lang>`). Empty when the block carries
 * no language (legacy blocks export bare fences, never a guess). Shared
 * with html.ts so both exporters agree. */
export function codeBlockLanguage(node: MarkdownNodeLike): string {
  try {
    const xml = node?.xml as Element | undefined;
    if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return '';
    const marks = (xml as Element).getElementsByTagName('w:bookmarkStart');
    for (let i = 0; i < marks.length; i++) {
      const name = (marks[i] as Element).getAttribute('w:name') || '';
      if (name.startsWith('_AstraCode_')) {
        const lang = name.slice('_AstraCode_'.length);
        if (/^[A-Za-z0-9_+#-]{1,24}$/.test(lang)) return lang;
      }
    }
  } catch { /* blocks without a language export bare */ }
  return '';
}

export interface MarkdownExportOptions {
  imageMode?: 'embed' | 'file';
  assetPrefix?: string;
}

export interface MarkdownImageAsset {
  name: string;
  mime: string;
  data: Buffer;
  alt: string;
}

function normalizeAssetPrefix(prefix?: string): string {
  const raw = (prefix ?? 'media').trim();
  if (!raw) {
    throw new Error('E_VALIDATION: asset prefix is empty. Remedy: pass a relative folder name such as media inside the export folder.');
  }
  if (raw.includes('\0') || /%00/i.test(raw) || raw.includes('\n') || raw.includes('\r')) {
    throw new Error('E_VALIDATION: asset prefix is not usable. Remedy: pass a plain relative folder name such as media.');
  }
  const posix = raw.replace(/\\/g, '/');
  if (posix.startsWith('/') || /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(posix)) {
    throw new Error(`E_VALIDATION: asset prefix "${raw}" points outside the export folder. Remedy: pass a relative folder name such as media inside the export folder.`);
  }
  const segs = posix.split('/').filter((s) => s.length > 0);
  if (segs.length === 0) {
    throw new Error('E_VALIDATION: asset prefix is empty. Remedy: pass a relative folder name such as media inside the export folder.');
  }
  for (const seg of segs) {
    if (seg === '.' || seg === '..') {
      throw new Error(`E_VALIDATION: asset prefix "${raw}" escapes the export folder. Remedy: keep the asset folder inside the export folder with a relative path.`);
    }
    if (seg.startsWith('.')) {
      throw new Error(`E_VALIDATION: asset prefix "${raw}" uses a hidden folder. Remedy: use a visible folder name such as media.`);
    }
    if (!/^[A-Za-z0-9._-]+$/.test(seg)) {
      throw new Error(`E_VALIDATION: asset prefix "${raw}" carries unsafe characters. Remedy: use letters, digits, dot, dash, or underscore such as media.`);
    }
  }
  return segs.join('/');
}

function mimeToImageExt(mime: string): string {
  const lower = (mime || '').toLowerCase();
  if (lower === 'image/jpeg') return 'jpg';
  if (lower === 'image/gif') return 'gif';
  if (lower === 'image/webp') return 'webp';
  if (lower === 'image/bmp') return 'bmp';
  if (lower === 'image/svg+xml') return 'svg';
  if (lower === 'image/tiff') return 'tif';
  return 'png';
}

function sanitizeImageAlt(alt: string, fallback: string): string {
  const flat = (alt || '').replace(/\r\n|\r|\n/g, ' ').replace(/[[\]]/g, '').trim();
  return flat || fallback;
}

function collectImageEntries(ast: DocumentAST): Array<{ node: Node; alt: string; mime: string; data: Buffer }> {
  const out: Array<{ node: Node; alt: string; mime: string; data: Buffer }> = [];
  for (const section of ast.sections) {
    for (const node of section.nodes) {
      if (node.type !== 'image' && !node.metadata?.hasImage) continue;
      const img = resolveNodeImage(ast, node);
      if (!img || !img.base64) continue;
      out.push({ node, alt: sanitizeImageAlt(img.alt, 'Image'), mime: img.mime || 'image/png', data: img.data });
    }
  }
  return out;
}

function buildImageNameMap(ast: DocumentAST, prefix: string): Map<string, string> {
  const entries = collectImageEntries(ast);
  const map = new Map<string, string>();
  entries.forEach((entry, idx) => {
    const num = String(idx + 1).padStart(4, '0');
    const name = `${prefix}/image-${num}.${mimeToImageExt(entry.mime)}`;
    map.set(entry.node.id, name);
  });
  return map;
}

export function collectMarkdownImageAssets(ast: DocumentAST, assetPrefix?: string): MarkdownImageAsset[] {
  const prefix = normalizeAssetPrefix(assetPrefix);
  return collectImageEntries(ast).map((entry, idx) => {
    const num = String(idx + 1).padStart(4, '0');
    return {
      name: `${prefix}/image-${num}.${mimeToImageExt(entry.mime)}`,
      mime: entry.mime,
      data: entry.data,
      alt: entry.alt,
    };
  });
}

/** True when a paragraph holds code across line breaks (w:br): either the
 * paragraph carries a code named style, or every text-bearing run is
 * monospace. Plain multi-line paragraphs (sibling inserts) fail
 * this gate and keep their normal rendering. Single-line code is excluded
 * by the caller (inline spans stay unchanged). */
function isCodeBlockParagraph(node: MarkdownNodeLike): boolean {
  const xml = node?.xml as Element | undefined;
  const metaStyle = String(metaRecordOf(node)?.['namedStyle'] || metaRecordOf(node)?.['style'] || '').toLowerCase();
  let pStyleVal = '';
  if (xml && typeof (xml as unknown as Element).getElementsByTagName === 'function') {
    const pStyleEl = (xml as Element).getElementsByTagName('w:pStyle')[0];
    if (pStyleEl) pStyleVal = ((pStyleEl as Element).getAttribute('w:val') || '').toLowerCase();
  }
  const named = metaStyle || pStyleVal;
  if (named === 'code' || named === 'monospace' || named === 'inlinecode') return true;
  if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return false;
  const runs = Array.from((xml as Element).getElementsByTagName('w:r'));
  let textRuns = 0;
  for (const r of runs) {
    // Accepted view: runs inside deletions carry no display text.
    let ancestor = (r as unknown as Element).parentNode as Element | null;
    let inDel = false;
    while (ancestor && ancestor !== (xml as Element)) {
      if (ancestor.tagName === 'w:del') { inDel = true; break; }
      ancestor = ancestor.parentNode as Element | null;
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

/** True when a paragraph carries the block box the code insert path writes
 * (a fill, gray by default or the style guide's code background, plus a border): marks a fenced-code block that may have no
 * language bookmark and no named style. Plain inline-code paragraphs never
 * match — their monospace runs sit in a bare paragraph. */
function hasCodeBlockBox(node: MarkdownNodeLike): boolean {
  const xml = node?.xml as Element | undefined;
  if (!xml || typeof (xml as unknown as Element).getElementsByTagName !== 'function') return false;
  const pPr = (xml as Element).getElementsByTagName('w:pPr')[0] as Element | undefined;
  if (!pPr) return false;
  const shd = pPr.getElementsByTagName('w:shd')[0] as Element | undefined;
  const fill = (shd?.getAttribute('w:fill') || '').toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(fill)) return false;
  const pBdr = pPr.getElementsByTagName('w:pBdr')[0] as Element | undefined;
  if (!pBdr) return false;
  const bottoms = pBdr.getElementsByTagName('w:bottom');
  if (bottoms.length === 0) return false;
  const val = (bottoms[0] as Element).getAttribute('w:val');
  return !!val && val !== 'none' && val !== 'nil';
}

function isDivider(node: MarkdownNodeLike, text?: string): boolean {
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
  if (metaRecordOf(node)?.['isDivider']) {
    return true;
  }
  const dividerXml = node.xml as Element | undefined;
  if (dividerXml && typeof (dividerXml as unknown as Element).getElementsByTagName === 'function') {
    const pBdrs = (dividerXml as Element).getElementsByTagName('w:pBdr');
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

function getAlertKind(style: string, meta?: NodeMetadata): { isAlert: boolean; alertType: string | null } {
  const explicit = (meta?.alertType || meta?.calloutType || meta?.callout?.type || '').toUpperCase();
  if (explicit) {
    if (['NOTE', 'TIP', 'IMPORTANT', 'WARNING', 'CAUTION'].includes(explicit)) {
      return { isAlert: true, alertType: explicit };
    }
  }
  if (meta?.callout) {
    const t = String(meta.callout.theme || '').toLowerCase();
    if (t === 'amber' || t.includes('warn')) return { isAlert: true, alertType: 'WARNING' };
    if (t === 'rose' || t.includes('caut')) return { isAlert: true, alertType: 'CAUTION' };
    if (t === 'emerald' || t.includes('tip')) return { isAlert: true, alertType: 'TIP' };
    if (meta.callout.title && /warning|caution/i.test(meta.callout.title)) return { isAlert: true, alertType: 'WARNING' };
    return { isAlert: true, alertType: 'NOTE' };
  }
  const s = style.toLowerCase();
  if (s.startsWith('heading') || s.startsWith('h1') || s.startsWith('h2') || s.startsWith('h3') || s === 'title' || s === 'tocheading') {
    return { isAlert: false, alertType: null };
  }
  if (s.includes('warning')) return { isAlert: true, alertType: 'WARNING' };
  if (s.includes('caution')) return { isAlert: true, alertType: 'CAUTION' };
  if (s.includes('important')) return { isAlert: true, alertType: 'IMPORTANT' };
  if (s.includes('tip')) return { isAlert: true, alertType: 'TIP' };
  if (s.includes('note') || s.includes('callout')) return { isAlert: true, alertType: 'NOTE' };
  if (s.includes('quote')) return { isAlert: true, alertType: null };
  return { isAlert: false, alertType: null };
}

function extractFootnoteBodies(parser?: DocxParser, ast?: DocumentAST): Map<string, string> {
  // Keys are namespaced `fn:<id>` (footnotes) vs `en:<id>` (endnotes) so a
  // footnote and an endnote sharing one numeric id never overwrite each other.
  const bodies = new Map<string, string>();

  // 1. ast.footnotes / ast.endnotes if available
  const ingestNoteMap = (src: unknown, prefix: 'fn' | 'en'): void => {
    if (!src) return;
    if (src instanceof Map) {
      for (const [k, v] of (src as Map<unknown, unknown>).entries()) {
        const noteText: unknown = typeof v === 'string' ? v : ((v as { text?: unknown } | null | undefined)?.text || String(v ?? ''));
        if (noteText) bodies.set(`${prefix}:${String(k)}`, String(noteText).trim());
      }
    } else if (Array.isArray(src)) {
      for (const item of src) {
        const rec = item as { id?: unknown; text?: unknown } | null | undefined;
        if (rec && rec.id != null) {
          bodies.set(`${prefix}:${String(rec.id)}`, String(rec.text || '').trim());
        }
      }
    } else if (typeof src === 'object') {
      for (const [k, v] of Object.entries(src as Record<string, unknown>)) {
        const noteText: unknown = typeof v === 'string' ? v : ((v as { text?: unknown } | null | undefined)?.text || String(v ?? ''));
        if (noteText) bodies.set(`${prefix}:${String(k)}`, String(noteText).trim());
      }
    }
  };
  if (ast) {
    const astNotes = ast as unknown as { footnotes?: unknown; endnotes?: unknown };
    ingestNoteMap(astNotes.footnotes, 'fn');
    ingestNoteMap(astNotes.endnotes, 'en');
  }

  // 2. parser if available
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

          // Extract paragraph text
          const paras = Array.from(el.getElementsByTagName('w:p'));
          const pTexts: string[] = [];
          if (paras.length > 0) {
            for (const p of paras) {
              const runs = Array.from(p.getElementsByTagName('w:r'));
              const runParts: string[] = [];
              for (const r of runs) {
                // Ignore footnoteRef / endnoteRef runs
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
      if (parser && typeof parser.getZipAdapter === 'function') {
        const adapter = parser.getZipAdapter();
        footnotesXml = adapter.readTextOrNull('word/footnotes.xml');
        endnotesXml = adapter.readTextOrNull('word/endnotes.xml');
      } else {
        const zip = parserZipOf(parser);
        if (zip && typeof zip.getEntry === 'function') {
          const fnEntry = zip.getEntry('word/footnotes.xml');
          if (fnEntry) footnotesXml = fnEntry.getData().toString('utf8');
          const enEntry = zip.getEntry('word/endnotes.xml');
          if (enEntry) endnotesXml = enEntry.getData().toString('utf8');
        }
      }
    } catch {
      /* ignore zip read error */
    }

    if (footnotesXml) parseNoteElements(footnotesXml, 'w:footnote', 'fn');
    if (endnotesXml) parseNoteElements(endnotesXml, 'w:endnote', 'en');
  }

  return bodies;
}

/**
 * L-3: Render table node to Markdown string, placing caption before or after table
 * based on node.metadata?.captionPosition ('above' | 'below').
 */
export function renderTableMarkdown(node: MarkdownNodeLike): string {
  const metaRec = metaRecordOf(node);
  const captionRaw = node.caption || metaRec['caption'] || metaRec['captionText'];
  let caption: string | undefined = typeof captionRaw === 'string' ? captionRaw : undefined;
  const nodeXmlEl = node.xml as Element | undefined;
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
  const captionPositionRaw = node.captionPosition || metaRec['captionPosition'] || metaRec['captionPlacement'] || 'above';
  const captionPosition: string = typeof captionPositionRaw === 'string' ? captionPositionRaw : 'above';

  let headers: string[] = [];
  let rows: string[][] = [];
  const metaHeaders: unknown = metaRec['headers'];
  const nodeHeaders: unknown = node.headers;
  if (Array.isArray(metaHeaders)) headers = metaHeaders.map(String);
  else if (Array.isArray(nodeHeaders)) headers = nodeHeaders.map(String);
  const metaRows: unknown = metaRec['rows'];
  const nodeRows: unknown = node.rows;
  const rowsSrc: unknown = Array.isArray(metaRows) && metaRows.length > 0 ? metaRows
    : (Array.isArray(nodeRows) && (nodeRows.length === 0 || Array.isArray((nodeRows as unknown[])[0])) ? nodeRows : []);
  if (Array.isArray(rowsSrc)) rows = (rowsSrc as unknown[]).map((r) => (Array.isArray(r) ? r.map(String) : [String(r ?? '')]));

  if (headers.length === 0 && rows.length === 0 && Array.isArray(node.rows)) {
    const extractedRows: string[][] = [];
    for (const r of node.rows as unknown[]) {
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

  if (headers.length === 0 && rows.length === 0 && nodeXmlEl) {
    const trs = Array.from((nodeXmlEl as Element).getElementsByTagName('w:tr')) as Element[];
    const tableData: string[][] = [];
    for (const tr of trs) {
      const tcs = Array.from(tr.getElementsByTagName('w:tc')) as Element[];
      const rowCells: string[] = [];
      for (const tc of tcs) {
        const textEls = Array.from(tc.getElementsByTagName('w:t'));
        const cellText = textEls.map((t: Element) => t.textContent || '').join(' ').trim().replace(/\|/g, '\\|');
        rowCells.push(cellText);
      }
      if (rowCells.length > 0) {
        tableData.push(rowCells);
      }
    }
    if (tableData.length > 0) {
      headers = tableData[0];
      rows = tableData.slice(1);
    }
  }

  const tableLines: string[] = [];
  if (headers.length > 0 || rows.length > 0) {
    const headerRow = headers.length > 0 ? headers : rows[0] || [];
    const dataRows = headers.length > 0 ? rows : rows.slice(1);
    tableLines.push(`| ${headerRow.join(' | ')} |`);
    tableLines.push(`| ${headerRow.map(() => '---').join(' | ')} |`);
    for (const row of dataRows) {
      tableLines.push(`| ${row.join(' | ')} |`);
    }
  }

  const parts: string[] = [];
  if (caption && captionPosition !== 'below') {
    parts.push(caption);
  }
  if (tableLines.length > 0) {
    parts.push(tableLines.join('\n'));
  }
  if (caption && captionPosition === 'below') {
    parts.push(caption);
  }
  return parts.join('\n');
}

export function convertAstToMarkdown(ast: DocumentAST, parser?: DocxParser, options?: MarkdownExportOptions): string;
export function convertAstToMarkdown(ast: DocumentAST, options?: MarkdownExportOptions): string;
export function convertAstToMarkdown(
  ast: DocumentAST,
  parserOrOptions?: DocxParser | MarkdownExportOptions,
  maybeOptions?: MarkdownExportOptions
): string {
  let parser: DocxParser | undefined;
  let options: MarkdownExportOptions | undefined;
  if (
    parserOrOptions &&
    typeof parserOrOptions === 'object' &&
    ('imageMode' in (parserOrOptions as Record<string, unknown>) ||
      'assetPrefix' in (parserOrOptions as Record<string, unknown>))
  ) {
    options = parserOrOptions as MarkdownExportOptions;
  } else {
    parser = parserOrOptions as DocxParser | undefined;
    options = maybeOptions;
  }
  const imageMode = options?.imageMode ?? 'embed';
  if (imageMode !== 'embed' && imageMode !== 'file') {
    throw new Error(
      `E_VALIDATION: image mode "${String((options as MarkdownExportOptions)?.imageMode)}" is not supported. Remedy: pass imageMode embed for data URLs or file for folder references.`
    );
  }
  let imageNameMap: Map<string, string> | null = null;
  if (imageMode === 'file') {
    const prefix = normalizeAssetPrefix(options?.assetPrefix);
    imageNameMap = buildImageNameMap(ast, prefix);
  }
  const lines: string[] = [];
  const footnoteOrder: string[] = [];
  const footnoteBodies = extractFootnoteBodies(parser, ast);

  for (const section of ast.sections) {
    if (section.heading) {
      const level = Math.min(section.heading.level || 1, 6);
      const prefix = '#'.repeat(level);
      lines.push(`${prefix} ${section.heading.text}\n`);
    }

    for (let ni = 0; ni < section.nodes.length; ni++) {
      const node = section.nodes[ni];
      // Runtime nodes may carry extended types beyond the AST union
      // (divider/callout/quote/hr); compare through a string view.
      const nodeType = node.type as string;
      if ((node.type === 'paragraph' && !hasListNumbering(node)) || nodeType === 'divider' || nodeType === 'horizontalRule' || nodeType === 'hr' || nodeType === 'callout') {
        const para = mdParagraph(node, ast);
        for (const id of para.footnoteIds) {
          if (!footnoteOrder.includes(id)) footnoteOrder.push(id);
        }
        const pStyleEl = node.xml?.getElementsByTagName?.('w:pStyle')?.[0] ?? null;
        const pStyleVal = pStyleEl ? pStyleEl.getAttribute('w:val') || '' : '';
        const style = (node.metadata?.namedStyle || (metaRecordOf(node)['style'] as string | undefined) || pStyleVal).toLowerCase();
        let text = para.text || (node.metadata?.text || '').trim();

        if (isDivider(node, text)) {
          // A divider paragraph may still carry text — never drop it.
          const pureMarker = text === '---' || text === '***' || text === '___';
          if (text && !pureMarker) lines.push(`${text}\n`);
          lines.push('---\n');
        } else if (style.startsWith('heading') || style.startsWith('h1') || style.startsWith('h2') || style.startsWith('h3')) {
          const match = style.match(/\d+/);
          const level = match ? Math.min(parseInt(match[0], 10), 6) : 1;
          if (text) lines.push(`${'#'.repeat(level)} ${text}\n`);
        } else if (style === 'title' || style === 'tocheading') {
          if (text) lines.push(`# ${text}\n`);
        } else {
          const alert = getAlertKind(style, node.metadata);
          if (alert.isAlert || nodeType === 'callout' || nodeType === 'quote') {
            const alertTag = alert.alertType || (nodeType === 'callout' ? 'NOTE' : null);
            if (alertTag) {
              const textLines = text ? text.split('\n') : [];
              lines.push(`> [!${alertTag}]\n${textLines.map(l => `> ${l}`).join('\n')}\n`);
            } else {
              // Quote
              const textLines = text ? text.split('\n') : [];
              lines.push(`${textLines.map(l => `> ${l}`).join('\n')}\n`);
            }
          } else {
            const isCodeNamedStyle = style === 'code' || style === 'monospace' || style === 'inlinecode';
            if (text && text.includes('\n') && (isCodeNamedStyle || isCodeBlockParagraph(node))) {
              // Multi-line code: fenced block with literal content (raw mode
              // skips the inline code-span wrapping; footnote/comment markers
              // reuse the shared para ctx merged above).
              let raw = '';
              if (node.xml && typeof node.xml?.getElementsByTagName === 'function') {
                raw = mdChildrenText(node.xml as Element, { footnoteIds: [], commentIds: [] }, ast, true).replace(/[ \t]+\n/g, '\n').trim();
              }
              if (!raw) raw = (node.metadata?.text || '').trim();
              if (raw) {
                lines.push(`${mdFencedBlock(raw, codeBlockLanguage(node))}\n`);
              }
            } else {
              // Single-line fenced code: the block gate above only fires on
              // '\n', but a one-line fence is still a block. Rebuild from raw
              // text and emit a fence when the paragraph reads as code and
              // carries a block marker — the fence-language bookmark or the
              // block box the code insert path writes. Named-style-only and
              // plain inline-code paragraphs keep their spans below.
              const singleLang = codeBlockLanguage(node);
              const singleIsCode = isCodeNamedStyle || isCodeBlockParagraph(node);
              const singleIsBlock = singleLang !== '' || hasCodeBlockBox(node);
              if (text && !text.includes('\n') && singleIsCode && singleIsBlock) {
                let rawFence = '';
                if (node.xml && typeof node.xml?.getElementsByTagName === 'function') {
                  rawFence = mdChildrenText(node.xml as Element, { footnoteIds: [], commentIds: [] }, ast, true).replace(/[ \t]+\n/g, '\n').trim();
                }
                if (!rawFence) rawFence = (node.metadata?.text || '').trim();
                if (rawFence) {
                  lines.push(`${mdFencedBlock(rawFence, singleLang)}\n`);
                }
              } else {
                if (isCodeNamedStyle && text && !text.includes('\n')) {
                  // Single-line code style renders literal content once. The
                  // styled text above may already carry code spans from
                  // monospace runs, so rebuild from raw text to avoid nesting
                  // one span inside another.
                  let rawSingle = '';
                  if (node.xml && typeof node.xml?.getElementsByTagName === 'function') {
                    rawSingle = mdChildrenText(node.xml as Element, { footnoteIds: [], commentIds: [] }, ast, true).replace(/[ \t]+\n/g, '\n').trim();
                  }
                  if (!rawSingle) rawSingle = (node.metadata?.text || '').trim();
                  if (rawSingle) {
                    text = mdCodeSpan(rawSingle);
                  }
                }
                if (text) {
                  lines.push(`${text}\n`);
                }
              }
            }
          }
        }
        if (para.commentIds.length > 0) {
          lines.push(`> [Comment ${para.commentIds.join(', ')} on ${node.id}]\n`);
        }
      } else if (node.type === 'table') {
        if (node.metadata?.callout) {
          const calloutMeta = node.metadata.callout;
          // An explicit alert type always wins over theme/title inference.
          const explicitRaw = String(
            node.metadata?.alertType ||
            node.metadata?.calloutType ||
            calloutMeta.type ||
            calloutMeta.alertType ||
            ''
          ).trim().toUpperCase();
          const explicitSet = ['NOTE', 'TIP', 'IMPORTANT', 'WARNING', 'CAUTION'];
          const theme = String(calloutMeta.theme || '').toLowerCase();
          const title = calloutMeta.title ? String(calloutMeta.title).trim() : '';
          const alertType = explicitSet.includes(explicitRaw) ? explicitRaw
            : theme === 'amber' || theme.includes('warn') || /warning/i.test(title) ? 'WARNING'
            : theme === 'rose' || theme.includes('caut') || /caution/i.test(title) ? 'CAUTION'
            : theme === 'emerald' || theme.includes('tip') || /tip/i.test(title) ? 'TIP'
            : 'NOTE';

          let text = '';
          if (node.metadata?.text) {
            text = node.metadata.text;
            if (title && text.startsWith(`${title}: `)) {
              text = text.slice(title.length + 2).trim();
            }
          } else if (node.xml) {
            const paras = Array.from(node.xml.getElementsByTagName('w:p'));
            text = paras.map(p => mdParagraph({ xml: p }, ast).text).join('\n').trim();
          }

          const alertLines = [`> [!${alertType}]`];
          if (title) {
            alertLines.push(`> **${title}**`);
          }
          if (text) {
            for (const line of text.split('\n')) {
              alertLines.push(`> ${line}`);
            }
          }
          lines.push(`${alertLines.join('\n')}\n`);
          continue;
        }

        const like = node as MarkdownNodeLike;
        const likeRec = metaRecordOf(like);
        let caption: string | undefined = typeof like.caption === 'string' ? like.caption
          : typeof likeRec['caption'] === 'string' ? likeRec['caption']
          : typeof likeRec['captionText'] === 'string' ? likeRec['captionText']
          : undefined;
        if (!caption && node.xml) {
          const tblPr = node.xml.getElementsByTagName('w:tblPr')[0];
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
        const captionPositionRaw = like.captionPosition || likeRec['captionPosition'] || likeRec['captionPlacement'] || 'above';
        const captionPosition: string = typeof captionPositionRaw === 'string' ? captionPositionRaw : 'above';

        let headers: string[] = [];
        let rows: string[][] = [];
        const convMetaHeaders: unknown = likeRec['headers'];
        const convNodeHeaders: unknown = like.headers;
        if (Array.isArray(convMetaHeaders)) headers = (convMetaHeaders as unknown[]).map(String);
        else if (Array.isArray(convNodeHeaders)) headers = (convNodeHeaders as unknown[]).map(String);
        const convMetaRows: unknown = likeRec['rows'];
        const convNodeRows: unknown = like.rows;
        const convRowsSrc: unknown = Array.isArray(convMetaRows) && convMetaRows.length > 0 ? convMetaRows
          : (Array.isArray(convNodeRows) && (convNodeRows.length === 0 || Array.isArray((convNodeRows as unknown[])[0])) ? convNodeRows : []);
        if (Array.isArray(convRowsSrc)) rows = (convRowsSrc as unknown[]).map((r) => (Array.isArray(r) ? (r as unknown[]).map(String) : [String(r ?? '')]));

        if (headers.length === 0 && rows.length === 0 && Array.isArray(like.rows)) {
          const extractedRows: string[][] = [];
          for (const r of like.rows as unknown[]) {
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

        // If metadata headers/rows are empty, extract directly from live XML DOM
        if (headers.length === 0 && rows.length === 0 && node.xml) {
          const trs = Array.from(node.xml.getElementsByTagName('w:tr')) as Element[];
          const tableData: string[][] = [];
          for (const tr of trs) {
            const tcs = Array.from(tr.getElementsByTagName('w:tc')) as Element[];
            const rowCells: string[] = [];
            for (const tc of tcs) {
              const textEls = Array.from(tc.getElementsByTagName('w:t'));
              const cellText = textEls.map((t: Element) => t.textContent || '').join(' ').trim().replace(/\|/g, '\\|');
              rowCells.push(cellText);
            }
            if (rowCells.length > 0) {
              tableData.push(rowCells);
            }
          }
          if (tableData.length > 0) {
            headers = tableData[0];
            rows = tableData.slice(1);
          }
        }

        if (caption && captionPosition !== 'below') {
          lines.push(`${caption}\n`);
        }

        if (headers.length > 0 || rows.length > 0) {
          const headerRow = headers.length > 0 ? headers : rows[0] || [];
          const dataRows = headers.length > 0 ? rows : rows.slice(1);
          lines.push(`| ${headerRow.join(' | ')} |`);
          lines.push(`| ${headerRow.map(() => '---').join(' | ')} |`);
          for (const row of dataRows) {
            lines.push(`| ${row.join(' | ')} |`);
          }
          lines.push('\n');
        }

        if (caption && captionPosition === 'below') {
          lines.push(`${caption}\n`);
        }
      } else if (node.type === 'image' || node.metadata?.hasImage) {
        // Embedded images render as data URLs by default. File mode emits
        // a folder reference plus a matching asset entry from
        // collectMarkdownImageAssets with the same deterministic name.
        const img = resolveNodeImage(ast, node);
        const safeAlt = (img?.alt || node.metadata?.text || 'Image').replace(/[[\]]/g, '') || 'Image';
        if (img && img.base64) {
          if (imageMode === 'file' && imageNameMap) {
            const ref = imageNameMap.get(node.id);
            if (ref) {
              const fileAlt = sanitizeImageAlt(img.alt, 'Image');
              lines.push(`![${fileAlt}](${ref})\n`);
            } else {
              lines.push(`![${safeAlt}](data:${img.mime};base64,${img.base64})\n`);
            }
          } else {
            lines.push(`![${safeAlt}](data:${img.mime};base64,${img.base64})\n`);
          }
        } else {
          // OX-07: never emit bare `![alt]` (no URL) — fallback text instead.
          lines.push(`[Image: ${safeAlt}]\n`);
        }
      } else if (nodeType === 'chart' || nodeType === 'chart3d') {
        const meta = node.metadata;
        const chartRec = (meta.chartData ?? {}) as Record<string, unknown>;
        const chartTitle: unknown = meta.title || chartRec['title'] || '';
        const title = typeof chartTitle === 'string' && chartTitle ? chartTitle : 'Chart Data Analysis';
        const chartTypeRaw: unknown = meta.chartType || chartRec['chartType'] || '';
        const chartType = typeof chartTypeRaw === 'string' ? chartTypeRaw : '';
        const categoriesSrc: unknown = meta.categories ?? chartRec['categories'];
        const categories: string[] = Array.isArray(categoriesSrc)
          ? (categoriesSrc as unknown[]).map((c) => String(c ?? ''))
          : ['Category 1', 'Category 2'];
        const seriesSrc: unknown = meta.series ?? chartRec['series'];
        const series: Array<{ name: string; values: unknown[] }> = Array.isArray(seriesSrc)
          ? (seriesSrc as Array<{ name?: unknown; values?: unknown }>).map((s) => ({
              name: typeof s?.name === 'string' ? s.name : String(s?.name ?? ''),
              values: Array.isArray(s?.values) ? [...(s.values as unknown[])] : [],
            }))
          : [{ name: 'Series 1', values: [10, 20] }];

        lines.push(`### 📊 ${title}\n`);
        if (chartType) lines.push(`*Chart type: ${chartType}*\n`);
        const header = ['Category', ...series.map(s => s.name)];
        lines.push(`| ${header.map(mdEscapePipes).join(' | ')} |`);
        lines.push(`| ${header.map(() => '---').join(' | ')} |`);
        categories.forEach((cat, cIdx) => {
          const rowVals = [cat, ...series.map(s => String(s.values[cIdx] ?? 0))];
          lines.push(`| ${rowVals.map(mdEscapePipes).join(' | ')} |`);
        });
        lines.push('\n');
      } else if (node.type === 'list' || (node.type === 'paragraph' && hasListNumbering(node))) {
        const para = mdParagraph(node, ast);
        const text = para.text || (node.metadata?.text || '').trim();
        const isNumbered = isOrderedListItem(node, ast);
        lines.push(`${isNumbered ? '1.' : '-'} ${text}`);
        // A list run needs a blank line before the next non-list block:
        // without it a following table (or any block) lexes as list-item
        // content and the exported Markdown is not re-importable.
        const next = section.nodes[ni + 1] as typeof node | undefined;
        const nextIsList =
          !!next && (next.type === 'list' || (next.type === 'paragraph' && hasListNumbering(next)));
        if (!nextIsList) lines.push('');
        for (const id of para.footnoteIds) {
          if (!footnoteOrder.includes(id)) footnoteOrder.push(id);
        }
        if (para.commentIds.length > 0) {
          lines.push(`> [Comment ${para.commentIds.join(', ')} on ${node.id}]\n`);
        }
      } else if (nodeType === 'equation') {
        // E4: standalone equation node — display `$$` block (verbatim latex source).
        const latex = (node.metadata?.text || mdEquationLatex(node.xml as Element) || '').trim();
        lines.push(latex ? `$$\n${latex}\n$$` : `[Equation]\n`);
      } else if (nodeType === 'field') {
        // OX-07: field node — display text only, codes omitted.
        const para = mdParagraph(node, ast);
        const text = para.text || (node.metadata?.text || '').trim() || '[Field]';
        lines.push(`${text}\n`);
        for (const id of para.footnoteIds) {
          if (!footnoteOrder.includes(id)) footnoteOrder.push(id);
        }
      } else {
        // OX-07 zero-silent-drop fallback.
        const text = (node.metadata?.text || '').trim();
        if (text) {
          lines.push(`${text}\n`);
        } else {
          lines.push(`[Unsupported node type: ${String(nodeType || 'unknown')}]\n`);
        }
      }
    }
  }

  // OX-07: headers/footers deduped; footnote defs close the document.
  const headerLines = mdZoneTexts(ast.headers).map(t => `*Header: ${t}*\n`);
  lines.unshift(...headerLines);
  for (const t of mdZoneTexts(ast.footers)) {
    lines.push(`*Footer: ${t}*\n`);
  }
  for (const key of footnoteOrder) {
    const sep = key.indexOf(':');
    const kind = sep >= 0 ? key.slice(0, sep) : 'fn';
    const id = sep >= 0 ? key.slice(sep + 1) : key;
    const label = kind === 'en' ? `en-${id}` : id;
    const realText = footnoteBodies.get(key);
    if (realText) {
      lines.push(`[^${label}]: ${realText}\n`);
    } else {
      lines.push(`[^${label}]: [Footnote ${label}: see document]\n`);
    }
  }

  return lines.join('\n');
}
