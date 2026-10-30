import type { Token } from 'marked';
import {
  ALLOWED_LINK_SCHEMES,
  AUTOLINK_PATTERN,
  CITATION_KEY_PATTERN,
  FOOTNOTE_REF_PATTERN,
  HTML_PATTERN,
  IMAGE_PATTERN,
  fail,
} from './constants';

export interface CitationDef {
  key: string;
  body: string;
  line: number;
}

export interface SpanSegment {
  text: string;
  bold: boolean;
  italic: boolean;
  code: boolean;
  linkUrl?: string;
  /** Inline math: `text` is LaTeX source, built into an `m:oMath` by compose. */
  math?: true;
  /** Markdown line of a math segment, for error messages. */
  line?: number;
}

export interface ParsedInline {
  plainText: string;
  segments: SpanSegment[];
  footnoteRefs: string[];
  citationKeys: string[];
}

export function failOnImageOrHtml(text: string, line: number, what: string): void {
  if (IMAGE_PATTERN.test(text)) {
    fail(line, `images are out of scope in ${what}`, 'attach images with the image insert op instead');
  }
  if (HTML_PATTERN.test(text) || AUTOLINK_PATTERN.test(text)) {
    fail(line, `raw HTML is out of scope in ${what}`, 'use plain Markdown blocks instead of HTML tags');
  }
}

const MARKED_ENTITIES: Record<string, string> = {
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#39;': "'",
};

/**
 * marked HTML-escapes inline text tokens (`'` -> `&#39;`, `&` -> `&amp;`, ...).
 * Run emission escapes again for XML, so the entities are decoded here, in one
 * pass, matching CommonMark (authored `&lt;` reads as `<`).
 */
export function unescapeMarkedText(text: string): string {
  return text.replace(/&(?:amp|lt|gt|quot|#39);/g, (m) => MARKED_ENTITIES[m]);
}

export function plainJoin(text: string): string {
  return text.replace(/\r\n|\r|\n/g, ' ').replace(/[ \t]+/g, ' ').trim();
}

/**
 * Newlines inside one paragraph read as spaces. Span segments keep their raw
 * text through compile, so this runs both when segments are built and right
 * before `w:t` runs emit: run output never carries a raw line break.
 */
export function nlToSpace(text: string): string {
  return text.replace(/\r\n|\r|\n/g, ' ');
}

export interface RefStrip {
  text: string;
  refs: string[];
}

export function stripRefs(text: string): RefStrip {
  const refs: string[] = [];
  const found = text.match(FOOTNOTE_REF_PATTERN);
  if (found) {
    for (const marker of found) {
      refs.push(marker.slice(2, -1));
    }
  }
  const cleaned = text.replace(/\s*\[\^[^\]]+\]\s*/g, ' ').replace(/[ \t]+/g, ' ').trim();
  return { text: cleaned, refs };
}

export function validateLinkUrl(href: string, line: number): string {
  const trimmed = href.trim();
  if (trimmed.length === 0) {
    fail(line, 'link has an empty address', 'write [text](https://example.com) with a real address');
  }
  if (trimmed.includes('\n') || trimmed.includes('\r') || trimmed.includes('\0')) {
    fail(line, 'link address carries a line break', 'keep the address on one line');
  }
  let scheme = '';
  let hasScheme = false;
  const schemeMatch = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(trimmed);
  if (schemeMatch) {
    scheme = `${schemeMatch[1].toLowerCase()}:`;
    hasScheme = true;
  } else {
    try {
      const parsed = new URL(trimmed);
      scheme = parsed.protocol.toLowerCase();
      hasScheme = true;
    } catch {
      scheme = '';
      hasScheme = false;
    }
  }
  if (hasScheme && !ALLOWED_LINK_SCHEMES.has(scheme)) {
    fail(line, `link scheme '${scheme}' is not allowed`, 'use http, https, mailto, ftp, tel, sms, or a relative address');
  }
  if (trimmed.length > 2000) {
    fail(line, 'link address is too long', 'keep addresses under 2000 characters');
  }
  return trimmed;
}

export type InlineToken = Token & {
  raw: string;
  text?: string;
  href?: string;
  tokens?: InlineToken[];
};

export function parseInlineTokens(
  inlineTokens: readonly InlineToken[] | undefined,
  line: number,
  what: string,
  citationDefs: Map<string, CitationDef>,
  citationOrder: string[],
  citationNum: Map<string, number>,
  usedCitations: Set<string>
): ParsedInline {
  const segments: SpanSegment[] = [];
  const footnoteRefs: string[] = [];
  const citationKeys: string[] = [];

  const pushTextWithMarkers = (
    rawText: string,
    flags: { bold: boolean; italic: boolean; code: boolean; linkUrl?: string },
    isCodeLiteral: boolean
  ): void => {
    if (rawText.length === 0) return;
    if (isCodeLiteral) {
      segments.push({ text: nlToSpace(rawText), bold: flags.bold, italic: flags.italic, code: true, linkUrl: flags.linkUrl });
      return;
    }
    const cursor = 0;
    const combined = /(\[\^([^\]]+)\])|(\[@([^\]]+)\])/g;
    let match: RegExpExecArray | null;
    let cleaned = '';
    void cleaned;
    const first = true;
    const flush = (piece: string): void => {
      if (piece.length === 0) return;
      if (first) {
        cleaned += piece;
      } else {
        cleaned += piece;
      }
    };
    void flush;
    const pieces: Array<{ kind: 'text' | 'footnote' | 'citation'; text?: string; label?: string; key?: string }> = [];
    let lastIndex = 0;
    combined.lastIndex = 0;
    while ((match = combined.exec(rawText)) !== null) {
      if (match.index > lastIndex) {
        pieces.push({ kind: 'text', text: rawText.slice(lastIndex, match.index) });
      }
      if (match[1] !== undefined) {
        pieces.push({ kind: 'footnote', label: match[2] });
      } else if (match[3] !== undefined) {
        pieces.push({ kind: 'citation', key: match[4] });
      }
      lastIndex = match.index + match[0].length;
    }
    if (lastIndex < rawText.length) {
      pieces.push({ kind: 'text', text: rawText.slice(lastIndex) });
    }
    if (pieces.length === 0) return;
    let segText = '';
    const commitSeg = (): void => {
      if (segText.length > 0) {
        segments.push({ text: segText, bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl });
        segText = '';
      }
    };
    void commitSeg;
    for (const piece of pieces) {
      if (piece.kind === 'text') {
        segText += piece.text as string;
      } else if (piece.kind === 'footnote') {
        const label = piece.label as string;
        if (label.trim().length === 0) {
          fail(line, 'empty footnote marker', 'write [^label] with a label');
        }
        footnoteRefs.push(label);
        // Marker strips out; surrounding spaces stay as they are.
      } else if (piece.kind === 'citation') {
        const rawKey = (piece.key as string).trim();
        if (!CITATION_KEY_PATTERN.test(rawKey)) {
          fail(line, `citation key [@${piece.key}] is malformed`, 'use [@key] with letters, digits, :, _, - and one key per bracket');
        }
        if (!citationDefs.has(rawKey)) {
          fail(line, `undefined citation [@${rawKey}]`, 'add a matching [@key]: entry definition');
        }
        usedCitations.add(rawKey);
        if (!citationNum.has(rawKey)) {
          citationOrder.push(rawKey);
          citationNum.set(rawKey, citationOrder.length);
        }
        const num = citationNum.get(rawKey) as number;
        citationKeys.push(rawKey);
        segText += `[${num}]`;
      }
    }
    if (segText.length > 0) {
      segments.push({ text: nlToSpace(segText), bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl });
    }
    void cursor;
  };

  const walk = (
    tokens: readonly InlineToken[],
    flags: { bold: boolean; italic: boolean; linkUrl?: string }
  ): void => {
    for (const tok of tokens) {
      const t = tok as InlineToken & { title?: unknown };
      switch (t.type) {
        case 'text':
        case 'escape':
          pushTextWithMarkers(unescapeMarkedText(t.text ?? ''), { bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl }, false);
          break;
        case 'strong':
          walk((t.tokens ?? []) as InlineToken[], { bold: true, italic: flags.italic, linkUrl: flags.linkUrl });
          break;
        case 'em':
          walk((t.tokens ?? []) as InlineToken[], { bold: flags.bold, italic: true, linkUrl: flags.linkUrl });
          break;
        case 'codespan':
          pushTextWithMarkers(unescapeMarkedText(t.text ?? ''), { bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl }, true);
          break;
        case 'link': {
          const raw = t.raw ?? '';
          if (raw.startsWith('<')) {
            fail(line, 'raw HTML is out of scope in ' + what, 'write links as [text](https://example.com) instead of <...>');
          }
          const footLink = raw.match(/^\[\^([^\]]+)\]$/);
          if (footLink) {
            const label = footLink[1];
            if (label.trim().length === 0) {
              fail(line, 'empty footnote marker', 'write [^label] with a label');
            }
            footnoteRefs.push(label);
            break;
          }
          const citeLink = raw.match(/^\[@([^\]]+)\]$/);
          if (citeLink) {
            const rawKey = citeLink[1].trim();
            if (!CITATION_KEY_PATTERN.test(rawKey)) {
              fail(line, `citation key [@${citeLink[1]}] is malformed`, 'use [@key] with letters, digits, :, _, - and one key per bracket');
            }
            if (!citationDefs.has(rawKey)) {
              fail(line, `undefined citation [@${rawKey}]`, 'add a matching [@key]: entry definition');
            }
            usedCitations.add(rawKey);
            if (!citationNum.has(rawKey)) {
              citationOrder.push(rawKey);
              citationNum.set(rawKey, citationOrder.length);
            }
            const num = citationNum.get(rawKey) as number;
            citationKeys.push(rawKey);
            segments.push({ text: `[${num}]`, bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl });
            break;
          }
          const href = validateLinkUrl(String(t.href ?? ''), line);
          if (flags.linkUrl !== undefined) {
            fail(line, 'nested links are out of scope', 'keep one link per span without links inside links');
          }
          walk((t.tokens ?? []) as InlineToken[], { bold: flags.bold, italic: flags.italic, linkUrl: href });
          break;
        }
        case 'parenMath':
        case 'dollarMath':
          if (flags.linkUrl !== undefined) {
            fail(line, 'inline math inside a link is out of scope', 'move the math outside the link text');
          }
          segments.push({ text: t.text ?? '', bold: false, italic: false, code: false, math: true, line });
          break;
        case 'image':
          fail(line, `images are out of scope in ${what}`, 'put images on their own block line instead of inside text');
          break;
        case 'html':
          fail(line, `raw HTML is out of scope in ${what}`, 'use plain Markdown blocks instead of HTML tags');
          break;
        case 'br':
          segments.push({ text: ' ', bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl });
          break;
        case 'del':
        case 'codespan2':
          segments.push({ text: t.raw ?? '', bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl });
          break;
        default: {
          const maybeTokens = (t as { tokens?: InlineToken[] }).tokens;
          if (Array.isArray(maybeTokens) && maybeTokens.length > 0) {
            walk(maybeTokens, flags);
          } else if (typeof t.text === 'string' && t.text.length > 0) {
            pushTextWithMarkers(unescapeMarkedText(t.text), { bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl }, false);
          } else if (typeof t.raw === 'string' && t.raw.length > 0) {
            segments.push({ text: t.raw, bold: flags.bold, italic: flags.italic, code: false, linkUrl: flags.linkUrl });
          }
          break;
        }
      }
    }
  };

  walk((inlineTokens ?? []) as InlineToken[], { bold: false, italic: false, linkUrl: undefined });
  const plainText = plainJoin(segments.map((s) => (s.math ? `\\(${s.text}\\)` : s.text)).join(''));
  return { plainText, segments, footnoteRefs, citationKeys };
}

export function segmentsNeedSplit(segments: SpanSegment[]): boolean {
  if (segments.length === 0) return false;
  if (segments.some((s) => s.math)) return true;
  if (segments.length > 1) {
    return segments.some((s) => s.bold || s.italic || s.code || s.linkUrl !== undefined);
  }
  const only = segments[0];
  return only.bold || only.italic || only.code || only.linkUrl !== undefined;
}
