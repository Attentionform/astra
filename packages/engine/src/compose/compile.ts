import type { Token } from 'marked';
import { lexMarkdown, MAX_BLOCK_MATH } from './math-extensions';
import type { Operation } from '../operation-schema';
import { toError } from '../errors';
import type { StyleGuideCalloutTag, StyleGuideConfig } from '../style-guide/types';
import { createDefaultStyleGuide } from '../style-guide/defaults';
import { containsSvgActiveContent, containsXxeAdjacent } from '../ooxml/xxe-guard';
import {
  decodeImagePixels,
  detectImageKind,
  fitImageEmus,
  readChartAsset,
  readImageAsset,
  resolveAllowedDir,
  resolveAssetPath,
} from './assets';
import {
  ALERT_MARKER_PATTERN,
  ALERT_TITLE_PATTERN,
  ALERT_TYPE_TO_THEME,
  AUTOLINK_PATTERN,
  CHART_FENCE_LANGUAGES,
  CHART_JSON_SUFFIX,
  CITATION_DEF_PATTERN,
  CITATION_KEY_PATTERN,
  CITATION_REF_PATTERN,
  COMPOSE_MARKDOWN_MAX_BYTES,
  COMPOSE_TABLE_HEADER_FILL,
  CONTAINER_FENCE_PATTERN,
  DATA_URL_BASE64_CHARS,
  DATA_URL_IMAGE_PATTERN,
  DATA_URL_MAX_BYTES,
  FOOTNOTE_DEF_PATTERN,
  FOOTNOTE_REF_PATTERN,
  HEADING_ATTR_PATTERN,
  HTML_PATTERN,
  IMAGE_PATTERN,
  fail,
  isEndnoteLabel,
} from './constants';
import type { CitationDef, InlineToken, SpanSegment } from './text-spans';
import { failOnImageOrHtml, parseInlineTokens, plainJoin, segmentsNeedSplit } from './text-spans';

interface FootnoteDef {
  label: string;
  body: string;
  line: number;
}

export interface CompileOptions {
  /** Directory that image and chart path references resolve against. Absent = path references fail closed. */
  allowedDir?: string;
  /**
   * Active style guide. Skins the two compose paths that carry no Markdown
   * styling of their own: table header fills (guide `tables.headerFill`)
   * and alert callout themes (guide `callouts.defaultTheme`). Markdown
   * itself has no per-table or per-alert theme syntax, so a guided compose
   * always prefers the guide. Omitted = the built-in "Default Corporate"
   * guide (navy header, bold white header text, shaded rows; alert colours
   * stay per tag); `'legacy'` = the old unskinned look (grey header).
   * Explicit op params still win downstream per the global precedence
   * invariant.
   */
  styleGuide?: StyleGuideConfig | 'legacy';
  /** Treat `$...$` as inline math (default off: prices and shell text can look like math). */
  dollarMath?: boolean;
}

/**
 * Guided table header fill: normalized 6-char uppercase guide
 * `tables.headerFill`, or undefined when no guide (or no usable fill) is
 * active. Undefined keeps the legacy built-in fill, byte-identical.
 */
/**
 * The guide a compose run uses. A caller-supplied guide wins; `'legacy'`
 * means no guide at all; omitted means the built-in default, minus its
 * single alert theme (Markdown alerts keep their per-tag colours).
 */
export function resolveComposeStyleGuide(option: StyleGuideConfig | 'legacy' | undefined): StyleGuideConfig | undefined {
  if (option === 'legacy') return undefined;
  if (option !== undefined) return option;
  const guide = createDefaultStyleGuide();
  if (guide.callouts) delete guide.callouts.defaultTheme;
  return guide;
}

function guidedTableHeaderFill(styleGuide?: StyleGuideConfig): string | undefined {
  return guidedHex(styleGuide?.tables?.headerFill);
}

/** Guide `tables.alternatingFill` as 6-char uppercase hex, or undefined (no guide, or no usable fill). */
function guidedAlternatingFill(styleGuide?: StyleGuideConfig): string | undefined {
  return guidedHex(styleGuide?.tables?.alternatingFill);
}

function guidedHex(raw: unknown): string | undefined {
  if (typeof raw !== 'string') return undefined;
  const norm = raw.replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(norm) ? norm : undefined;
}

function lineAt(src: string, offset: number): number {
  let line = 1;
  for (let i = 0; i < offset && i < src.length; i++) {
    if (src.charCodeAt(i) === 10) line++;
  }
  return line;
}

/** Exact start line per block token via sequential raw matching. */
function buildLineIndex(src: string, tokens: readonly Token[]): Map<Token, number> {
  const out = new Map<Token, number>();
  let cursor = 0;
  for (const token of tokens) {
    const raw = (token as { raw?: unknown }).raw;
    if (typeof raw !== 'string' || raw.length === 0) {
      out.set(token, lineAt(src, cursor));
      continue;
    }
    const found = src.indexOf(raw, cursor);
    const at = found === -1 ? cursor : found;
    out.set(token, lineAt(src, at));
    cursor = at + raw.length;
  }
  return out;
}

/** Line of a sub-excerpt inside a block that starts at `blockLine`. */

function subLine(blockRaw: string, blockLine: number, excerpt: string): number {
  const found = blockRaw.indexOf(excerpt);
  if (found === -1) return blockLine;
  return blockLine + lineAt(blockRaw, found) - 1;
}

type BlockToken = Token & {
  raw: string;
  text?: string;
  depth?: number;
  ordered?: boolean;
  tokens?: InlineToken[];
  items?: Array<{
    raw: string;
    text: string;
    task: boolean;
    tokens?: Token[];
  }>;
  header?: Array<{ text: string; tokens?: InlineToken[] }>;
  rows?: Array<Array<{ text: string; tokens?: InlineToken[] }>>;
  lang?: string;
};

/**
 * Compiles Markdown to validated engine ops (pure except asset reads: no
 * session IO, deterministic ids). Throws `E_VALIDATION` with a 1-based line
 * number on any out-of-scope or malformed input, before any caller can write.
 * Inputs larger than `COMPOSE_MARKDOWN_MAX_BYTES` UTF-8 bytes throw `E_LIMIT`
 * before the lexer runs.
 */
export function compileMarkdown(markdown: string, options?: CompileOptions): Operation[] {
  return compileMarkdownDetailed(markdown, options).operations;
}

/** `compileMarkdown` plus the advisories compile found (input that was kept but probably not meant). */
export function compileMarkdownDetailed(
  markdown: string,
  options?: CompileOptions
): { operations: Operation[]; warnings: string[] } {
  const warnings: string[] = [];
  if (typeof markdown !== 'string') {
    throw toError('E_VALIDATION', 'line 1: compose input must be a string', 'pass the Markdown source as a string');
  }
  const rawBytes = Buffer.byteLength(markdown, 'utf8');
  if (rawBytes > COMPOSE_MARKDOWN_MAX_BYTES) {
    throw toError(
      'E_LIMIT',
      `markdown source ${rawBytes} bytes exceeds compose ceiling ${COMPOSE_MARKDOWN_MAX_BYTES} bytes`,
      `split the manuscript into smaller parts under ${COMPOSE_MARKDOWN_MAX_BYTES} bytes or reduce content, then retry`
    );
  }
  const src = markdown.replace(/\r\n|\r/g, '\n');
  if (!/[^\s]/.test(src)) {
    fail(1, 'no content blocks found', 'pass at least one heading, paragraph, list, table, divider, or code block');
  }

  let allowedDir: string | undefined;
  if (options?.allowedDir !== undefined) {
    allowedDir = resolveAllowedDir(options.allowedDir, 1);
  }
  const styleGuide = resolveComposeStyleGuide(options?.styleGuide);

  const srcLines = src.split('\n');
  let inFence = false;
  const defs = new Map<string, FootnoteDef>();
  const citationDefs = new Map<string, CitationDef>();
  for (let i = 0; i < srcLines.length; i++) {
    const lineText = srcLines[i];
    if (/^\s{0,3}(`{3,}|~{3,})/.test(lineText)) {
      inFence = !inFence;
      continue;
    }
    if (inFence) continue;
    const footMatch = lineText.match(FOOTNOTE_DEF_PATTERN);
    if (footMatch) {
      const label = footMatch[1];
      const body = (footMatch[2] ?? '').trim();
      const line = i + 1;
      if (defs.has(label)) {
        fail(line, `duplicate footnote definition [^${label}]`, 'keep one definition per label');
      }
      if (body.length === 0) {
        fail(line, `footnote definition [^${label}] has an empty body`, 'write the note text after the colon on the same line');
      }
      if (IMAGE_PATTERN.test(body) || HTML_PATTERN.test(body) || AUTOLINK_PATTERN.test(body)) {
        fail(line, `footnote definition [^${label}] carries image or HTML markup`, 'keep note bodies plain text');
      }
      if (FOOTNOTE_REF_PATTERN.test(body)) {
        FOOTNOTE_REF_PATTERN.lastIndex = 0;
        fail(line, `footnote definition [^${label}] references another note`, 'flatten nested notes into plain text');
      }
      FOOTNOTE_REF_PATTERN.lastIndex = 0;
      if (CITATION_REF_PATTERN.test(body)) {
        CITATION_REF_PATTERN.lastIndex = 0;
        fail(line, `footnote definition [^${label}] carries a citation`, 'move citations into body text');
      }
      CITATION_REF_PATTERN.lastIndex = 0;
      defs.set(label, { label, body, line });
      continue;
    }
    const citeMatch = lineText.match(CITATION_DEF_PATTERN);
    if (citeMatch) {
      const rawKey = citeMatch[1].trim();
      const body = (citeMatch[2] ?? '').trim();
      const line = i + 1;
      if (!CITATION_KEY_PATTERN.test(rawKey)) {
        fail(line, `citation definition [@${citeMatch[1]}] has a malformed key`, 'use [@key] with letters, digits, :, _, -');
      }
      if (citationDefs.has(rawKey)) {
        fail(line, `duplicate citation definition [@${rawKey}]`, 'keep one definition per key');
      }
      if (defs.has(rawKey)) {
        fail(line, `citation definition [@${rawKey}] collides with a note label`, 'rename one of them');
      }
      if (body.length === 0) {
        fail(line, `citation definition [@${rawKey}] has an empty body`, 'write the entry text after the colon on the same line');
      }
      if (IMAGE_PATTERN.test(body) || HTML_PATTERN.test(body) || AUTOLINK_PATTERN.test(body)) {
        fail(line, `citation definition [@${rawKey}] carries image or HTML markup`, 'keep entries plain text');
      }
      if (FOOTNOTE_REF_PATTERN.test(body)) {
        FOOTNOTE_REF_PATTERN.lastIndex = 0;
        fail(line, `citation definition [@${rawKey}] references a note`, 'move note references into body text');
      }
      FOOTNOTE_REF_PATTERN.lastIndex = 0;
      if (CITATION_REF_PATTERN.test(body)) {
        CITATION_REF_PATTERN.lastIndex = 0;
        fail(line, `citation definition [@${rawKey}] references another citation`, 'flatten nested citations into plain text');
      }
      CITATION_REF_PATTERN.lastIndex = 0;
      citationDefs.set(rawKey, { key: rawKey, body, line });
    }
  }

  let tokens: Token[];
  try {
    tokens = lexMarkdown(src, { dollarMath: options?.dollarMath });
  } catch (err: unknown) {
    const detail = err instanceof Error ? err.message : String(err);
    throw toError('E_VALIDATION', `line 1: Markdown lexer rejected the input (${detail})`, 'simplify the failing block and retry');
  }
  const lineOf = buildLineIndex(src, tokens);

  // Raw table-intent scan. The lexer silently normalizes ragged rows to the
  // header width (dropping cells), so row widths are checked here on source
  // lines before anything emits. Any pipe-bearing line directly above a
  // dash-separator row reads as table intent unless it sits inside a list,
  // code, quote, or HTML block (literal there) or already lexed cleanly.
  const LITERAL_RANGE_KINDS = new Set(['list', 'code', 'blockquote', 'html']);
  const tokenRanges = tokens.map((entry) => {
    const entryRaw = (entry as { raw?: unknown }).raw;
    const start = lineOf.get(entry) ?? 1;
    const span =
      typeof entryRaw === 'string' && entryRaw.length > 0
        ? entryRaw.replace(/\n+$/, '').split('\n').length
        : 1;
    return { start, end: start + span - 1, kind: entry.type };
  });
  const inLiteralRange = (lineNo: number): boolean =>
    tokenRanges.some(
      (range) => LITERAL_RANGE_KINDS.has(range.kind) && range.start <= lineNo && lineNo <= range.end
    );
  const PIPE_GUARD = '\uE000';
  const splitTableCells = (row: string): string[] => {
    const guarded = row.replace(/\\\|/g, PIPE_GUARD);
    const parts = guarded.split('|');
    if (parts.length > 0 && /^\s*$/.test(parts[0])) parts.shift();
    if (parts.length > 0 && /^\s*$/.test(parts[parts.length - 1])) parts.pop();
    return parts.map((cell) => cell.split(PIPE_GUARD).join('|'));
  };
  const isTableDelimRow = (row: string): boolean => {
    if (!row.includes('|')) return false;
    const cells = splitTableCells(row);
    return cells.length > 0 && cells.every((cell) => /^\s*:?-+:?\s*$/.test(cell));
  };
  for (let i = 0; i + 1 < srcLines.length; i++) {
    const headerRow = srcLines[i];
    const delimRow = srcLines[i + 1];
    if (!headerRow.includes('|') || /^\s*$/.test(headerRow) || !isTableDelimRow(delimRow)) continue;
    const headerLine = i + 1;
    if (inLiteralRange(headerLine)) {
      i += 1;
      continue;
    }
    const headerCells = splitTableCells(headerRow);
    const delimCells = splitTableCells(delimRow);
    if (delimCells.length !== headerCells.length) {
      fail(
        headerLine,
        `table separator row has ${delimCells.length} cells, header has ${headerCells.length}`,
        'give the separator row exactly as many dash cells as the header'
      );
    }
    let j = i + 2;
    let bodyRow = 0;
    while (j < srcLines.length && srcLines[j].includes('|') && !/^\s*$/.test(srcLines[j])) {
      bodyRow += 1;
      const count = splitTableCells(srcLines[j]).length;
      if (count !== headerCells.length) {
        fail(
          headerLine,
          `table row ${bodyRow} has ${count} cells, header has ${headerCells.length}`,
          'give every row exactly as many cells as the header'
        );
      }
      j += 1;
    }
    i = j - 1;
  }

  const ops: Operation[] = [];
  const usedDefs = new Set<string>();
  const usedCitations = new Set<string>();
  const citationOrder: string[] = [];
  const citationNum = new Map<string, number>();
  const pendingNotes: Array<{ anchor: string; label: string; line: number; endnote: boolean }> = [];
  let seq = 0;
  let lastInsertId: string | null = null;

  const nextId = (): string => {
    seq += 1;
    return `md_${String(seq).padStart(4, '0')}`;
  };

  const insertAnchor = (): { anchor: string; position: 'before' | 'after' } => {
    if (lastInsertId === null) return { anchor: 'end', position: 'after' };
    return { anchor: lastInsertId, position: 'after' };
  };

  const pushInsertParagraph = (text: string, spans?: SpanSegment[]): string => {
    const id = nextId();
    const anchor = insertAnchor();
    const op: Record<string, unknown> = {
      action: 'insert',
      type: 'paragraph',
      anchor: anchor.anchor,
      position: anchor.position,
      text,
      id,
    };
    if (spans !== undefined && segmentsNeedSplit(spans)) {
      op['params'] = { mdSpans: spans };
    }
    ops.push(op as Operation);
    lastInsertId = id;
    return id;
  };

  const trackRefs = (anchorId: string, refs: string[], line: number): void => {
    for (const label of refs) {
      usedDefs.add(label);
      pendingNotes.push({ anchor: anchorId, label, line, endnote: isEndnoteLabel(label) });
    }
  };

  const checkCitationDefsForNotes = (body: string, line: number, owner: string): void => {
    if (FOOTNOTE_REF_PATTERN.test(body)) {
      FOOTNOTE_REF_PATTERN.lastIndex = 0;
      fail(line, `${owner} references a note`, 'move note references into body text');
    }
    FOOTNOTE_REF_PATTERN.lastIndex = 0;
  };

  const replaceCitationsInPlainCell = (cellText: string, line: number): string => {
    return cellText.replace(CITATION_REF_PATTERN, (_whole, rawKey: string) => {
      const key = String(rawKey).trim();
      if (!CITATION_KEY_PATTERN.test(key)) {
        fail(line, `citation key [@${rawKey}] is malformed`, 'use [@key] with letters, digits, :, _, - and one key per bracket');
      }
      if (!citationDefs.has(key)) {
        fail(line, `undefined citation [@${key}]`, 'add a matching [@key]: entry definition');
      }
      usedCitations.add(key);
      if (!citationNum.has(key)) {
        citationOrder.push(key);
        citationNum.set(key, citationOrder.length);
      }
      return `[${citationNum.get(key) as number}]`;
    });
  };

  const failOnSpanTokensInCell = (cellTokens: readonly InlineToken[] | undefined, line: number): void => {
    const walkCell = (list: readonly InlineToken[]): void => {
      for (const tok of list) {
        const t = tok as InlineToken;
        if (t.type === 'strong' || t.type === 'em' || t.type === 'codespan') {
          fail(line, 'rich text inside tables is out of scope', 'move bold, italic, or code into a body paragraph');
        }
        if (t.type === 'link') {
          const raw = (t as { raw?: unknown }).raw;
          if (typeof raw === 'string' && raw.startsWith('<')) {
            fail(line, 'raw HTML is out of scope in a table', 'use plain Markdown blocks instead of HTML tags');
          }
          fail(line, 'links inside tables are out of scope', 'move the link into a body paragraph');
        }
        if (t.type === 'image') {
          fail(line, 'images are out of scope in a table', 'put images on their own block line instead');
        }
        if (t.type === 'html') {
          fail(line, 'raw HTML is out of scope in a table', 'use plain Markdown blocks instead of HTML tags');
        }
        if (t.type === 'parenMath' || t.type === 'dollarMath') {
          fail(line, 'inline math inside tables is out of scope', 'move the math into a body paragraph');
        }
        const nested = (t as { tokens?: InlineToken[] }).tokens;
        if (Array.isArray(nested) && nested.length > 0) walkCell(nested);
      }
    };
    if (cellTokens) walkCell(cellTokens);
  };

  const handleBlockImage = (
    href: string,
    alt: string,
    line: number
  ): { id: string; kind: 'image' | 'chart' } => {
    const target = href.trim();
    if (target.length === 0) {
      fail(line, 'image reference has an empty path', 'point the image at a file inside the allowed directory');
    }
    if (/^data:/i.test(target)) {
      // Inline data-URL image: bytes decode in memory through the same
      // magic-byte and size gates as file assets, then embed identically.
      const dataMatch = DATA_URL_IMAGE_PATTERN.exec(target);
      if (!dataMatch) {
        fail(line, 'image data URL is not a supported image', 'use data:image/png;base64,... (png, jpg, gif, webp, bmp, tiff, or svg) with base64 bytes');
      }
      const compact = target.slice((dataMatch as RegExpExecArray)[0].length).replace(/\s+/g, '');
      if (compact.length === 0) {
        fail(line, 'image data URL carries no bytes', 'paste base64 image bytes after data:image/...;base64,');
      }
      if (!DATA_URL_BASE64_CHARS.test(compact)) {
        fail(line, 'image data URL carries invalid base64', 'paste clean base64 bytes without other characters');
      }
      let inlineBuf: Buffer;
      try {
        inlineBuf = Buffer.from(compact, 'base64');
      } catch {
        fail(line, 'image data URL cannot be decoded', 'paste clean base64 image bytes');
      }
      inlineBuf = inlineBuf as Buffer;
      if (inlineBuf.length === 0) {
        fail(line, 'image data URL decoded to zero bytes', 'paste base64 bytes of a real image');
      }
      if (inlineBuf.length > DATA_URL_MAX_BYTES) {
        fail(line, 'image data URL is too large', 'use a smaller image (decoded bytes stay under fifteen megabytes)');
      }
      const inlineKind = detectImageKind(inlineBuf);
      if (!inlineKind) {
        fail(line, 'image data URL bytes are not a known image', 'use a real png, jpg, gif, webp, bmp, tiff, or svg payload');
      }
      if (inlineKind === 'svg') {
        const svgText = inlineBuf.toString('utf8');
        if (/<!DOCTYPE|<!ENTITY/i.test(svgText) || /<script/i.test(svgText) || containsXxeAdjacent(svgText) || containsSvgActiveContent(svgText)) {
          fail(line, 'image SVG carries active or entity markup', 'use a plain SVG without doctype, entities, script, event attributes, or javascript: hrefs (plain <style> blocks pass ingest and are stripped with report on sanitize)');
        }
      }
      const inlineMime =
        inlineKind === 'jpeg' ? 'image/jpeg' : inlineKind === 'svg' ? 'image/svg+xml' : inlineKind === 'tiff' ? 'image/tiff' : `image/${inlineKind}`;
      const inlineFitted = fitImageEmus(decodeImagePixels(inlineBuf, inlineKind));
      const inlineId = nextId();
      const inlineAnchor = lastInsertId === null ? 'end' : lastInsertId;
      ops.push({
        action: 'insert',
        type: 'image',
        anchor: inlineAnchor,
        position: 'after',
        id: inlineId,
        fileData: `data:${inlineMime};base64,${inlineBuf.toString('base64')}`,
        ...(inlineFitted ? { width: inlineFitted.width, height: inlineFitted.height } : {}),
        ...(alt.trim().length > 0 ? { altText: alt.trim() } : {}),
      } as Operation);
      lastInsertId = inlineId;
      return { id: inlineId, kind: 'image' };
    }
    if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(target)) {
      fail(line, 'image reference points outside the allowed directory', 'keep referenced files inside the allowed directory and use a relative path');
    }
    if (allowedDir === undefined) {
      fail(line, 'image needs the folder holding the picture file', 'pass allowedDir pointing at the folder with the picture (CLI --allowed-dir), or move the picture into that folder');
    }
    const isChart = target.toLowerCase().endsWith(CHART_JSON_SUFFIX);
    const absPath = resolveAssetPath(allowedDir as string, target, line, isChart ? 'chart' : 'image');
    const id = nextId();
    const anchorId = lastInsertId === null ? 'end' : lastInsertId;
    if (isChart) {
      const asset = readChartAsset(absPath, line, allowedDir as string);
      const title = asset.title ?? (alt.trim().length > 0 ? alt.trim() : 'Chart');
      ops.push({
        action: 'chart',
        operation: 'insert',
        anchor: anchorId,
        id,
        chartType: asset.chartType,
        title,
        ...(asset.categories ? { categories: asset.categories } : {}),
        ...(asset.series ? { series: asset.series } : {}),
      } as Operation);
      lastInsertId = id;
      return { id, kind: 'chart' };
    }
    const asset = readImageAsset(absPath, line, allowedDir as string);
    const anchor = lastInsertId === null ? { anchor: 'end', position: 'after' as const } : { anchor: lastInsertId, position: 'after' as const };
    ops.push({
      action: 'insert',
      type: 'image',
      anchor: anchor.anchor,
      position: anchor.position,
      id,
      fileData: asset.dataUrl,
      ...(asset.width !== undefined && asset.height !== undefined ? { width: asset.width, height: asset.height } : {}),
      ...(alt.trim().length > 0 ? { altText: alt.trim() } : {}),
    } as Operation);
    lastInsertId = id;
    return { id, kind: 'image' };
  };

  const failOnSpanTokensInAlert = (alertTokens: readonly InlineToken[] | undefined, line: number): void => {
    const walkAlert = (list: readonly InlineToken[]): void => {
      for (const tok of list) {
        const t = tok as InlineToken;
        if (t.type === 'strong' || t.type === 'em' || t.type === 'codespan') {
          fail(line, 'rich text inside alerts is out of scope', 'move bold, italic, or code into a body paragraph');
        }
        if (t.type === 'link') {
          const raw = (t as { raw?: unknown }).raw;
          if (typeof raw === 'string' && raw.startsWith('<')) {
            fail(line, 'raw HTML is out of scope in an alert', 'use plain Markdown blocks instead of HTML tags');
          }
          fail(line, 'links inside alerts are out of scope', 'move the link into a body paragraph');
        }
        if (t.type === 'image') {
          fail(line, 'images are out of scope in an alert', 'put images on their own block line instead');
        }
        if (t.type === 'html') {
          fail(line, 'raw HTML is out of scope in an alert', 'use plain Markdown blocks instead of HTML tags');
        }
        if (t.type === 'parenMath' || t.type === 'dollarMath') {
          fail(line, 'inline math inside alerts is out of scope', 'move the math into a body paragraph');
        }
        const nested = (t as { tokens?: InlineToken[] }).tokens;
        if (Array.isArray(nested) && nested.length > 0) walkAlert(nested);
      }
    };
    if (alertTokens) walkAlert(alertTokens);
  };

  /**
   * Compiles a `> [!TYPE]` alert blockquote to a real `callout` insert.
   * Every other quote shape fails closed as a general blockquote. Bodies
   * stay plain text (mirrors the table-cell policy: bold, italic, code,
   * links, images, and note markers fail instead of dropping formatting),
   * while citation keys still resolve to `[n]` text. An optional bold-only
   * first line becomes the callout title. Footnote markers fail: the callout
   * node id is minted at execution time, so there is no compile-time anchor
   * for a note (same reason as table cells). Theme: Markdown carries no
   * explicit theme syntax, so every alert tag is generic — the active guide
   * default wins when one is set, else the legacy per-tag map.
   */
  const handleAlertBlockquote = (token: BlockToken, line: number): void => {
    const inner = (token.tokens ?? []) as Token[];
    if (inner.length === 0) {
      fail(line, 'blockquotes are out of scope', 'use a plain paragraph');
    }
    for (const child of inner) {
      if (child.type !== 'paragraph') {
        fail(line, `alert content carries ${child.type} blocks, out of scope`, 'keep alert bodies plain text');
      }
    }
    const firstLines = (((inner[0] as BlockToken).text ?? '').split('\n'));
    const marker = ALERT_MARKER_PATTERN.exec(firstLines[0].trim());
    if (!marker) {
      fail(line, 'blockquotes are out of scope', 'use a plain paragraph');
    }
    const alertType = (marker[1] as string).toUpperCase();
    let rest = firstLines.slice(1).join('\n');
    const sameLineBody = ((marker[2] ?? '') as string).trim();
    if (sameLineBody.length > 0) {
      rest = rest.length > 0 ? `${sameLineBody}\n${rest}` : sameLineBody;
    }
    for (let i = 1; i < inner.length; i++) {
      rest += (rest.length > 0 ? '\n\n' : '') + (((inner[i] as BlockToken).text ?? ''));
    }
    const restLines = rest.split('\n');
    let title: string | undefined;
    let bodySrc = rest;
    if (restLines.length > 0 && ALERT_TITLE_PATTERN.test(restLines[0].trim())) {
      const innerTitle = (ALERT_TITLE_PATTERN.exec(restLines[0].trim())?.[1] ?? '').trim();
      if (innerTitle.length === 0) {
        fail(line, `alert [!${alertType}] title is empty`, 'write text between the ** markers or drop the title line');
      }
      if (innerTitle.length > 1000) {
        fail(line, `alert [!${alertType}] title is too long`, 'keep the title under 1000 characters');
      }
      failOnImageOrHtml(innerTitle, line, 'an alert title');
      if (FOOTNOTE_REF_PATTERN.test(innerTitle)) {
        FOOTNOTE_REF_PATTERN.lastIndex = 0;
        fail(line, 'footnote markers inside alerts are out of scope', 'move the note reference into a body paragraph');
      }
      FOOTNOTE_REF_PATTERN.lastIndex = 0;
      if (CITATION_REF_PATTERN.test(innerTitle)) {
        CITATION_REF_PATTERN.lastIndex = 0;
        fail(line, 'citation keys inside alert titles are out of scope', 'move the citation into the alert body');
      }
      CITATION_REF_PATTERN.lastIndex = 0;
      title = innerTitle;
      bodySrc = restLines.slice(1).join('\n');
    }
    for (const bodyLine of bodySrc.split('\n')) {
      if (CONTAINER_FENCE_PATTERN.test(bodyLine)) {
        fail(line, 'container fences (:::) are out of scope as custom styles', 'use plain Markdown blocks');
      }
      if (FOOTNOTE_DEF_PATTERN.test(bodyLine) || CITATION_DEF_PATTERN.test(bodyLine)) {
        fail(line, 'note and citation definitions must stand alone as their own block', 'move the definition out of the alert');
      }
    }
    if (!/[^\s]/.test(bodySrc)) {
      fail(line, `alert [!${alertType}] needs body text`, 'write the alert text on > lines after the marker');
    }
    let bodyTokens: Token[];
    try {
      bodyTokens = lexMarkdown(bodySrc, { dollarMath: options?.dollarMath });
    } catch (err: unknown) {
      const detail = err instanceof Error ? err.message : String(err);
      throw toError('E_VALIDATION', `line ${line}: alert body rejected (${detail})`, 'keep alert bodies plain text');
    }
    const cleanParas: string[] = [];
    for (const bodyToken of bodyTokens) {
      if (bodyToken.type === 'space') continue;
      if (bodyToken.type !== 'paragraph') {
        fail(line, `alert [!${alertType}] carries ${bodyToken.type} blocks, out of scope`, 'keep alert bodies plain text');
      }
      const para = bodyToken as BlockToken;
      const paraText = para.text ?? '';
      const paraLine = subLine(token.raw, line, (para.raw ?? '').split('\n')[0] ?? '');
      if (FOOTNOTE_REF_PATTERN.test(paraText)) {
        FOOTNOTE_REF_PATTERN.lastIndex = 0;
        fail(paraLine, 'footnote markers inside alerts are out of scope', 'move the note reference into a body paragraph');
      }
      FOOTNOTE_REF_PATTERN.lastIndex = 0;
      failOnImageOrHtml(paraText, paraLine, 'an alert');
      failOnSpanTokensInAlert(para.tokens, paraLine);
      cleanParas.push(plainJoin(replaceCitationsInPlainCell(paraText, paraLine)));
    }
    const text = plainJoin(cleanParas.join(' '));
    if (text.length === 0) {
      fail(line, `alert [!${alertType}] needs body text`, 'write the alert text on > lines after the marker');
    }
    const anchorId = lastInsertId === null ? 'end' : lastInsertId;
    // Theme precedence for composed alerts (per-tag themes live in the
    // guide file — Markdown source stays generic, every [!TAG] is plain
    // GitHub-flavored Markdown): a configured per-tag theme wins, else the
    // session guide default wins when one is active, else the legacy per-tag
    // map (guideless output stays byte-identical). An explicit callout-op
    // theme (batch path) still beats both downstream.
    const alertTag = alertType.toLowerCase() as StyleGuideCalloutTag;
    const tagThemeRaw = styleGuide?.callouts?.tagThemes?.[alertTag];
    const tagTheme = typeof tagThemeRaw === 'string' && tagThemeRaw.length > 0 ? tagThemeRaw : undefined;
    const guideThemeRaw = styleGuide?.callouts?.defaultTheme;
    const guideTheme = typeof guideThemeRaw === 'string' && guideThemeRaw.length > 0 ? guideThemeRaw : undefined;
    const op: Record<string, unknown> = {
      action: 'callout',
      anchor: anchorId,
      position: 'after',
      text,
      theme: tagTheme ?? guideTheme ?? ALERT_TYPE_TO_THEME[alertType] ?? 'indigo',
    };
    if (title !== undefined) op['title'] = title;
    ops.push(op as Operation);
    // Callout node ids mint at execution time, so no compile-time id exists
    // to chain onto: reset to the live end anchor for the next insert.
    lastInsertId = null;
  };

  // True right after a level-1 heading that opens the document, so an italic-only line under it can become the Subtitle.
  let underTitle = false;

  let lastBlockKind: 'none' | 'paragraph' | 'heading' | 'list' | 'table' | 'code' | 'image' | 'chart' | 'divider' | 'callout' | 'equation' = 'none';

  for (const rawToken of tokens) {
    const token = rawToken as BlockToken;
    const line = lineOf.get(rawToken) ?? 1;
    switch (token.type) {
      case 'space':
        break;

      case 'heading': {
        const depth = token.depth ?? 0;
        if (depth < 1 || depth > 3) {
          fail(line, `heading level ${depth} is out of scope (levels 1-3 only)`, 'use #, ##, or ###');
        }
        const text = token.text ?? '';
        if (HEADING_ATTR_PATTERN.test(text)) {
          fail(line, 'custom heading attributes are out of scope', 'drop the {#id}/{.class} suffix and use a plain heading');
        }
        failOnImageOrHtml(text, line, 'a heading');
        const inline = parseInlineTokens(token.tokens, line, 'a heading', citationDefs, citationOrder, citationNum, usedCitations);
        for (const key of inline.citationKeys) usedCitations.add(key);
        for (const label of inline.footnoteRefs) {
          if (!defs.has(label)) {
            fail(line, `undefined footnote reference [^${label}]`, 'add a matching [^label]: body definition');
          }
        }
        checkCitationDefsForNotes(text, line, 'a heading');
        const id = pushInsertParagraph(inline.plainText, inline.segments);
        trackRefs(id, inline.footnoteRefs, line);
        const styleName = depth === 1 ? 'Heading1' : depth === 2 ? 'Heading2' : 'Heading3';
        ops.push({ action: 'namedStyle', operation: 'apply', styleName, target: id } as Operation);
        underTitle = depth === 1 && lastBlockKind === 'none';
        if (lastBlockKind === 'table' && depth === 3) {
          ops.push({ action: 'style', target: id, spacingBefore: 240 } as Operation);
        }
        lastBlockKind = 'heading';
        break;
      }

      case 'paragraph': {
        const text = token.text ?? '';
        const rawLines = token.raw.replace(/\n+$/, '').split('\n');
        if (/^\s*(\$\$|\\\[)\s*$/.test(rawLines[0] ?? '')) {
          fail(
            line,
            `math block opened with ${/^\s*\$\$/.test(rawLines[0] ?? '') ? '$$' : '\\['} is never closed, or is longer than ${MAX_BLOCK_MATH} characters`,
            'close it with a line containing only $$ (or \\]), and keep the block under the length limit'
          );
        }
        if (rawLines.length > 0 && FOOTNOTE_DEF_PATTERN.test(rawLines[0])) {
          if (rawLines.length > 1 || !/^\s{0,3}\[\^[^\]]+\]:\s*\S/.test(token.raw)) {
            fail(line, 'footnote definition must stand alone as its own block', 'separate the definition from surrounding text with blank lines');
          }
          break;
        }
        if (rawLines.some((entry) => FOOTNOTE_DEF_PATTERN.test(entry))) {
          fail(line, 'footnote definition must stand alone as its own block', 'separate the definition from surrounding text with blank lines');
        }
        if (rawLines.length > 0 && CITATION_DEF_PATTERN.test(rawLines[0])) {
          if (rawLines.length > 1 || !/^\s{0,3}\[@[^\]]+\]:\s*\S/.test(token.raw)) {
            fail(line, 'citation definition must stand alone as its own block', 'separate the definition from surrounding text with blank lines');
          }
          break;
        }
        if (rawLines.some((entry) => CITATION_DEF_PATTERN.test(entry))) {
          fail(line, 'citation definition must stand alone as its own block', 'separate the definition from surrounding text with blank lines');
        }
        if (CONTAINER_FENCE_PATTERN.test(rawLines[0] ?? '')) {
          fail(line, 'container fences (:::) are out of scope as custom styles', 'use plain Markdown blocks');
        }
        const inlineTokens = (token.tokens ?? []) as InlineToken[];
        if (inlineTokens.length === 1 && inlineTokens[0].type === 'image') {
          const img = inlineTokens[0] as InlineToken & { href?: string };
          const imgRes = handleBlockImage(String(img.href ?? ''), String(img.text ?? ''), line);
          if (lastBlockKind === 'table') {
            ops.push({ action: 'style', target: imgRes.id, spacingBefore: 240 } as Operation);
          }
          lastBlockKind = imgRes.kind;
          break;
        }
        for (const it of inlineTokens) {
          if (it.type === 'image') {
            fail(line, 'images are out of scope in a paragraph', 'put images on their own block line instead of inside text');
          }
        }
        failOnImageOrHtml(text, line, 'a paragraph');
        const inline = parseInlineTokens(token.tokens, line, 'a paragraph', citationDefs, citationOrder, citationNum, usedCitations);
        for (const key of inline.citationKeys) usedCitations.add(key);
        for (const label of inline.footnoteRefs) {
          if (!defs.has(label)) {
            fail(line, `undefined footnote reference [^${label}]`, 'add a matching [^label]: body definition');
          }
        }
        const strayAlert = ALERT_MARKER_PATTERN.exec(rawLines[0]?.trim() ?? '');
        if (strayAlert) {
          warnings.push(
            `line ${line}: "[!${(strayAlert[1] as string).toUpperCase()}]" is not inside a blockquote, so it was written as plain text. Start each line with "> " to make it a callout.`
          );
        }
        // An italic-only line straight under the title is the document subtitle.
        // The Subtitle style already italicises, so the emphasis is dropped.
        const isSubtitle =
          underTitle &&
          lastBlockKind === 'heading' &&
          rawLines.length === 1 &&
          inlineTokens.length === 1 &&
          inlineTokens[0].type === 'em';
        const id = pushInsertParagraph(inline.plainText, isSubtitle ? undefined : inline.segments);
        trackRefs(id, inline.footnoteRefs, line);
        if (isSubtitle) {
          ops.push({ action: 'namedStyle', operation: 'apply', styleName: 'Subtitle', target: id } as Operation);
        }
        underTitle = false;
        if (lastBlockKind === 'table') {
          ops.push({ action: 'style', target: id, spacingBefore: 240 } as Operation);
        }
        lastBlockKind = 'paragraph';
        break;
      }

      case 'list': {
        const listType = token.ordered === true ? 'numbered' : 'bullet';
        const items = token.items ?? [];
        if (items.length === 0) {
          fail(line, 'empty list block', 'add at least one list item');
        }
        const itemIds: string[] = [];
        const itemBaseLine = line;
        for (const item of items) {
          const itemLine = subLine(token.raw, itemBaseLine, item.raw);
          if (item.task === true) {
            fail(itemLine, 'task-list items are out of scope', 'use a plain bullet or numbered item');
          }
          for (const child of item.tokens ?? []) {
            if (child.type === 'list') {
              const nestedRaw = (child as { raw?: unknown }).raw;
              const nestedLine =
                typeof nestedRaw === 'string' && nestedRaw
                  ? subLine(token.raw, itemBaseLine, nestedRaw)
                  : itemLine;
              fail(nestedLine, 'nested lists are out of scope', 'flatten nested items into the top-level list');
            }
            if (child.type !== 'text' && child.type !== 'paragraph' && child.type !== 'space') {
              fail(itemLine, `list items carry ${child.type} blocks, out of scope`, 'keep list items plain text');
            }
          }
          if (item.text.split('\n').some((entry) => FOOTNOTE_DEF_PATTERN.test(entry))) {
            fail(itemLine, 'footnote definition must stand alone as its own block', 'move the definition out of the list');
          }
          if (item.text.split('\n').some((entry) => CITATION_DEF_PATTERN.test(entry))) {
            fail(itemLine, 'citation definition must stand alone as its own block', 'move the definition out of the list');
          }
          failOnImageOrHtml(item.text, itemLine, 'a list item');
          const firstChild = (item.tokens ?? [])[0] as (Token & { tokens?: InlineToken[] }) | undefined;
          const inlineTokens = ((firstChild as { tokens?: InlineToken[] })?.tokens ?? []) as InlineToken[];
          if (inlineTokens.length === 1 && inlineTokens[0].type === 'image') {
            fail(itemLine, 'images are out of scope in a list item', 'put images on their own block line instead of inside a list');
          }
          for (const it of inlineTokens) {
            if (it.type === 'image') {
              fail(itemLine, 'images are out of scope in a list item', 'put images on their own block line instead of inside a list');
            }
          }
          const inline = parseInlineTokens(inlineTokens, itemLine, 'a list item', citationDefs, citationOrder, citationNum, usedCitations);
          for (const key of inline.citationKeys) usedCitations.add(key);
          for (const label of inline.footnoteRefs) {
            if (!defs.has(label)) {
              fail(itemLine, `undefined footnote reference [^${label}]`, 'add a matching [^label]: body definition');
            }
          }
          const id = pushInsertParagraph(inline.plainText, inline.segments);
          trackRefs(id, inline.footnoteRefs, itemLine);
          itemIds.push(id);
        }
        ops.push({ action: 'list', operation: 'create', target: itemIds, listType } as Operation);
        if (lastBlockKind === 'table' && itemIds.length > 0) {
          ops.push({ action: 'style', target: itemIds[0], spacingBefore: 240 } as Operation);
        }
        lastBlockKind = 'list';
        break;
      }

      case 'table': {
        const header = (token.header ?? []).map((cell) => cell.text ?? '');
        const rows = (token.rows ?? []).map((row) => row.map((cell) => cell.text ?? ''));
        if (header.length === 0) {
          fail(line, 'table has no header columns', 'write a header row plus a separator row (| --- |)');
        }
        for (let r = 0; r < rows.length; r++) {
          const cells = rows[r];
          if (cells.length !== header.length) {
            fail(
              line,
              `table row ${r + 1} has ${cells.length} cells, header has ${header.length}`,
              'give every row exactly as many cells as the header'
            );
          }
          for (const cell of cells) {
            if (FOOTNOTE_REF_PATTERN.test(cell)) {
              FOOTNOTE_REF_PATTERN.lastIndex = 0;
              fail(line + 2 + r, 'footnote markers inside tables are out of scope', 'move the note reference into a body paragraph');
            }
            FOOTNOTE_REF_PATTERN.lastIndex = 0;
          }
        }
        failOnImageOrHtml(token.raw, line, 'a table');
        const headerCells = (token.header ?? []) as Array<{ text: string; tokens?: InlineToken[] }>;
        const rowCells = (token.rows ?? []) as Array<Array<{ text: string; tokens?: InlineToken[] }>>;
        for (const cell of headerCells) failOnSpanTokensInCell(cell.tokens, line);
        for (const row of rowCells) for (const cell of row) failOnSpanTokensInCell(cell.tokens, line + 2);
        const cleanHeaders = header.map((h) => replaceCitationsInPlainCell(h, line));
        const cleanRows = rows.map((row) => row.map((cell) => replaceCitationsInPlainCell(cell, line)));
        for (const h of cleanHeaders) {
          if (FOOTNOTE_REF_PATTERN.test(h)) {
            FOOTNOTE_REF_PATTERN.lastIndex = 0;
            fail(line, 'footnote markers inside tables are out of scope', 'move the note reference into a body paragraph');
          }
          FOOTNOTE_REF_PATTERN.lastIndex = 0;
        }
        const id = nextId();
        const anchor = insertAnchor();
        ops.push({
          action: 'insert',
          type: 'table',
          anchor: anchor.anchor,
          position: anchor.position,
          headers: cleanHeaders,
          rows: cleanRows,
          id,
          cellPadding: { top: 120, bottom: 120, left: 160, right: 160 },
        } as Operation);
        lastInsertId = id;
        // Guided header fill (guide tables.headerFill; legacy D9D9D9
        // without a guide, byte-identical). Emitted directly: a bare table
        // insert leaves headers unshaded, so omitting the fill would lose
        // the header band instead of skinning it. Header text auto-paint
        // follows downstream via the guide contrast rule.
        const headerFill = guidedTableHeaderFill(styleGuide) ?? COMPOSE_TABLE_HEADER_FILL;
        for (let c = 0; c < header.length; c++) {
          ops.push({
            action: 'table',
            target: id,
            operation: 'setCellShading',
            rowIndex: 1,
            colIndex: c + 1,
            fill: headerFill,
          } as Operation);
        }
        // Guide alternating fill: shaded odd data rows, painted per cell (the
        // header keeps its own fill). Legacy composes carry no guide, so no stripes.
        const stripe = guidedAlternatingFill(styleGuide);
        if (stripe !== undefined && rows.length > 1) {
          ops.push({ action: 'table', target: id, operation: 'setShading', fill: stripe, zebra: true } as Operation);
        }
        lastBlockKind = 'table';
        break;
      }

      case 'blockMath': {
        const latex = String((token as { text?: unknown }).text ?? '').trim();
        if (latex.length === 0) {
          fail(line, 'empty math block', 'write the LaTeX between the $$ lines');
        }
        const id = nextId();
        const anchor = insertAnchor();
        ops.push({
          action: 'equation',
          operation: 'insert',
          anchor: anchor.anchor,
          position: anchor.position,
          latex,
          mode: 'latex',
          id,
        } as Operation);
        lastInsertId = id;
        if (lastBlockKind === 'table') {
          ops.push({ action: 'style', target: id, spacingBefore: 240 } as Operation);
        }
        lastBlockKind = 'equation';
        break;
      }
      case 'hr': {
        const id = nextId();
        const anchor = insertAnchor();
        ops.push({
          action: 'insert',
          type: 'divider',
          anchor: anchor.anchor,
          position: anchor.position,
          id,
        } as Operation);
        lastInsertId = id;
        if (lastBlockKind === 'table') {
          ops.push({ action: 'style', target: id, spacingBefore: 240 } as Operation);
        }
        lastBlockKind = 'divider';
        break;
      }

      case 'code': {
        const lang = (token.lang ?? '').trim();
        if (CHART_FENCE_LANGUAGES.has(lang.toLowerCase())) {
          fail(line, `fenced code tagged '${lang}' requests a chart`, 'describe the chart with a chart path reference instead');
        }
        const id = nextId();
        const anchor = insertAnchor();
        // Code blocks style at insert time (params.codeBlock): the insert
        // path applies the Code character style (Consolas/indigo/10pt) with
        // block shading, border, and indent, and records the fence language
        // in a bookmark so export rebuilds the fence with its tag.
        ops.push({
          action: 'insert',
          type: 'paragraph',
          anchor: anchor.anchor,
          position: anchor.position,
          text: token.text ?? '',
          id,
          params: { codeBlock: true, ...(lang ? { codeLanguage: lang } : {}) },
        } as Operation);
        lastInsertId = id;
        if (lastBlockKind === 'table') {
          ops.push({ action: 'style', target: id, spacingBefore: 240 } as Operation);
        }
        lastBlockKind = 'code';
        break;
      }

      case 'def': {
        const defLine = token.raw.replace(/\n+$/, '');
        if (FOOTNOTE_DEF_PATTERN.test(defLine) && !defLine.includes('\n')) {
          break;
        }
        if (CITATION_DEF_PATTERN.test(defLine) && !defLine.includes('\n')) {
          break;
        }
        fail(line, 'link definitions are out of scope', 'write links as [text](https://example.com) instead of a definition');
        break;
      }

      case 'html':
        fail(line, 'raw HTML blocks are out of scope', 'use plain Markdown blocks instead of HTML tags');
        break;

      case 'blockquote':
        handleAlertBlockquote(token, line);
        lastBlockKind = 'callout';
        break;

      case 'image': {
        const img = token as BlockToken & { href?: string };
        const href = String((img as { href?: unknown }).href ?? '');
        const alt = String(token.text ?? '');
        const imgRes = handleBlockImage(href, alt, line);
        if (lastBlockKind === 'table') {
          ops.push({ action: 'style', target: imgRes.id, spacingBefore: 240 } as Operation);
        }
        lastBlockKind = imgRes.kind;
        break;
      }

      default:
        fail(line, `block type '${token.type}' is out of scope`, 'use headings, paragraphs, lists, tables, dividers, code fences, or footnotes');
    }
  }

  for (const [label, def] of defs) {
    if (!usedDefs.has(label)) {
      if (isEndnoteLabel(label)) {
        fail(def.line, `unused endnote definition [^${label}] (no reference in text)`, 'reference it with [^label] or delete the definition');
      }
      fail(def.line, `unused footnote definition [^${label}] (no reference in text)`, 'reference it with [^label] or delete the definition');
    }
  }
  for (const [key, def] of citationDefs) {
    if (!usedCitations.has(key)) {
      fail(def.line, `unused citation definition [@${key}] (no reference in text)`, 'reference it with [@key] or delete the definition');
    }
  }

  for (const note of pendingNotes) {
    const def = defs.get(note.label);
    if (!def) {
      fail(note.line, `undefined footnote reference [^${note.label}]`, 'add a matching [^label]: body definition');
    }
    if (note.endnote) {
      ops.push({ action: 'endnote', anchor: note.anchor, text: (def as FootnoteDef).body } as Operation);
    } else {
      ops.push({ action: 'footnote', anchor: note.anchor, text: (def as FootnoteDef).body } as Operation);
    }
  }

  if (citationOrder.length > 0) {
    const bibHeadingId = nextId();
    const bibAnchor = insertAnchor();
    ops.push({
      action: 'insert',
      type: 'paragraph',
      anchor: bibAnchor.anchor,
      position: bibAnchor.position,
      text: 'References',
      id: bibHeadingId,
    } as Operation);
    lastInsertId = bibHeadingId;
    ops.push({ action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: bibHeadingId } as Operation);
    for (const key of citationOrder) {
      const def = citationDefs.get(key);
      if (!def) {
        fail(1, `undefined citation [@${key}]`, 'add a matching [@key]: entry definition');
      }
      const num = citationNum.get(key) as number;
      const entryId = nextId();
      const entryAnchor = insertAnchor();
      ops.push({
        action: 'insert',
        type: 'paragraph',
        anchor: entryAnchor.anchor,
        position: entryAnchor.position,
        text: `[${num}] ${(def as CitationDef).body}`,
        id: entryId,
      } as Operation);
      lastInsertId = entryId;
    }
  }

  if (ops.length === 0) {
    fail(1, 'no content blocks found', 'pass at least one heading, paragraph, list, table, divider, or code block');
  }
  return { operations: ops, warnings };
}
