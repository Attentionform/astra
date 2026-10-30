/**
 * Markdown composer: compiles a Markdown string to validated engine ops.
 *
 * Pure compiler except for asset reads: `compileMarkdown` maps Markdown blocks
 * to canonical operation JSON; `composeDocument` feeds that list through the
 * existing `executeBatch` path unchanged, so every gate (schema, quotas,
 * structural checks, atomic rollback) applies as usual. When image or chart
 * path references are present, `compileMarkdown` reads those files from disk
 * under an explicit allowed directory (otherwise it fails closed); all other
 * validation still throws before any caller can write.
 *
 * Supported input (anything else fails closed with a line number):
 * - `#` / `##` / `###` headings: paragraph insert + named-style apply
 *   (`Heading1`..`Heading3`). Deeper levels fail. Inline bold, italic, code,
 *   and links inside headings get run splitting (see inline notes below).
 * - Paragraphs separated by blank lines. Single newlines inside one
 *   paragraph join with a space (documented). Inline bold, italic, code, and
 *   links get run splitting.
 * - Flat bullet (`-`, `*`, `+`) and ordered (`1.`, `1)`) lists: one
 *   paragraph insert per item plus a single `list` create op carrying real
 *   `numPr` numbering. Ordered start numbers are ignored (documented).
 *   Nested lists, task items, and block content inside items fail. Inline
 *   spans inside items get run splitting.
 * - GitHub-style tables with a header separator row: `insert` table op
 *   (grid column counts, repeat header row) plus one `setCellShading` op per
 *   header cell (light gray fill, or the guide `tables.headerFill` when a
 *   style guide is active — Markdown carries no per-table styling, so the
 *   guide wins outright). Rows whose cell count differs from the header
 *   fail. Column alignment markers are accepted and ignored
 *   (documented). Footnote markers and images inside cells fail (cell text
 *   has no anchorable node of its own). Bold, italic, code, and link markup
 *   inside cells fail as well (cells carry plain strings only); move rich
 *   text into a body paragraph. Citation keys inside cells are allowed: they
 *   are plain text replacements (`[@key]` becomes `[n]`) that feed the
 *   bibliography without needing an anchor.
 * - `---` / `***` / `___` rules: real divider inserts. No page-break or
 *   section-break op is ever emitted, so break spam is impossible here.
 * - Fenced (or indented) code: paragraph insert with line breaks preserved
 *   (they land as `w:br`), the template Code character style on the runs
 *   (Consolas/indigo/10pt) with block shading, border, and indent on the
 *   paragraph, and the fence language in a bookmark so exporters rebuild
 *   the fence with its tag. Fence content stays fully literal: no footnote, endnote,
 *   citation, image, chart, or markup processing happens inside it. Fences
 *   tagged `chart`/`mermaid` fail (use a chart path reference instead, see
 *   below).
 * - Footnote references `[^label]` in paragraphs, headings, and list items,
 *   with `[^label]: body` definitions collected independently
 *   of position (use-before-define is fine). Markers are stripped from the
 *   inserted text; each note becomes a real `footnote` insert anchored to
 *   its paragraph node. Undefined references, unused definitions, and empty
 *   bodies fail. Notes inside notes fail. Markers may sit inside bold or
 *   italic runs; markers inside inline code stay literal.
 * - Endnote references `[^en-name]` with `[^en-name]: body` definitions use
 *   the same shape as footnotes but route through the existing `endnote`
 *   insert path. Any label starting with `en-` is an endnote; all other
 *   labels are footnotes. Rules match footnotes: use-before-define is fine,
 *   markers strip from text, undefined references, unused definitions, and
 *   empty bodies fail, notes inside notes fail, markers inside tables fail.
 * - Citation keys `[@key]` with `[@key]: entry` definitions. Keys hold
 *   letters, digits, and `:`, `_`, `-` (must start with a letter or digit).
 *   Each in-text key becomes `[n]` where `n` follows first appearance order,
 *   and the compiler appends a `References` heading plus one paragraph per
 *   entry (`[n] entry`) in that order. Entries are manual text: no formatter
 *   runs here, the definition body is used verbatim (documented manual
 *   pattern, no formatter engine). Undefined keys, unused definitions, empty
 *   bodies, and duplicate definitions fail. Definitions must stand alone as
 *   their own block. Keys inside fenced code and inline code stay literal.
 *   Keys inside footnote or endnote bodies fail (move the citation into body
 *   text). Multi-key brackets such as `[@a; @b]` fail: write one bracket per
 *   key.
 * - Image and chart path references as standalone block images:
 *   `![alt](path)` on its own paragraph. When the target ends with `.json`
 *   (any case) it is a chart definition file; otherwise it is an image file.
 *   Both require an explicit allowed directory (see options): the reference
 *   resolves against it and stays jailed inside it (symlinks canonicalized,
 *   outside paths fail). Images become real `image` inserts with file bytes
 *   embedded as a data URL and alt text preserved; charts become real
 *   `chart` inserts with type, title, categories, and series read from the
 *   JSON file. Inline base64 data-URL images (`![alt](data:image/png;base64,
 *   ...)`) need no allowed directory: the bytes decode in memory, validate
 *   against the same magic-byte and size gates (decoded output capped at
 *   fifteen megabytes, garbage fails closed), and embed through the same
 *   image insert path with fitted sizes. Image inserts carry fitted `width`/`height` (EMUs) decoded
 *   from the file's intrinsic pixel size at 96 DPI, capped to the page
 *   content box (6in wide, 9in tall) with the height scaled in proportion,
 *   so pictures keep their aspect ratio instead of stretching to the engine
 *   fallback box. Files whose pixel size cannot be decoded omit the size and
 *   fall back to the engine defaults; this fallback is the known sizing gap.
 *   Undecodable means any of: truncated headers (PNG under 24 bytes, GIF
 *   under 10, BMP under 26, WEBP under 16, TIFF under 8); BMP with an
 *   unsupported DIB size (neither 12 nor >= 40); JPEG with no start-of-frame
 *   marker (SOS or EOI hit first) or a truncated frame or segment length;
 *   WEBP that is not RIFF/WEBP, carries an unknown chunk (not VP8, VP8L, or
 *   VP8X), or truncates inside the chunk size fields; TIFF with a bad byte
 *   order or magic, a bad IFD offset or entry count (zero or over 200), a
 *   missing width or height tag, or width/height stored in an unsupported
 *   field type or count; SVG with no `svg` tag, no usable width, height, or
 *   viewBox (missing, unparsable units, non-positive numbers), or viewBox
 *   parts that are not four finite numbers; any decoded dimensions that are
 *   non-finite, non-positive, or over 100000 px; or fitted EMUs that are
 *   non-finite or non-positive. Inline images
 *   mixed with text fail (use a standalone block).
 *   Remote URLs as references fail (no network fetch; attach
 *   bytes via the allowed directory, or paste an inline data URL instead).
 *   Missing files, bad extensions, bad magic bytes, bad chart JSON, and paths
 *   outside the allowed directory fail with a line number. Without an allowed
 *   directory, any path reference fails closed (inline data URLs still work,
 *   since they carry their own bytes).
 * - Inline spans in paragraphs, headings, and list items: `**bold**`,
 *   `__bold__`, `*italic*`, `_italic_`, `` `code` ``, and `[text](url)`.
 *   Spans become multiple runs inside one paragraph (run splitting): bold
 *   sets `w:b`, italic sets `w:i`, code sets a monospace face, links become
 *   real `w:hyperlink` entries with external relationships plus the usual
 *   link color and underline. Combined spans (bold italic, bold link) keep
 *   both properties. Autolinks (`<https://...>`, `<user@...>`) still fail as
 *   raw HTML (write `[text](url)` instead). Bad link schemes (`javascript:`,
 *   `data:`, `vbscript:`, `file:`) fail. Images and raw HTML inside text
 *   fail. Unsupported inline markup stays literal text.
 *
 * Fail-closed with `E_VALIDATION` + line number: remote image URLs, raw HTML
 * (block or inline, including autolinks), custom-style syntax (`:::` fences,
 * `{#id}`/`{.class}` heading attributes), general blockquotes, deeper
 * headings, nested tables (only expressible via HTML, caught by that rule),
 * and any other unlisted block shape. Inputs larger than
 * `COMPOSE_MARKDOWN_MAX_BYTES` UTF-8 bytes fail closed with `E_LIMIT`
 * before the lexer runs (split the manuscript and retry). Alert blockquotes (`> [!NOTE]` and the
 * matching TIP, IMPORTANT, WARNING, CAUTION tags) are the one accepted quote
 * shape: they compile to real `callout` containers with plain-text bodies
 * (an optional bold-only first line becomes the callout title; bold, italic,
 * code, links, images, and note markers inside the body fail closed, like
 * table cells, so no formatting is silently dropped). Alert tags carry no
 * explicit theme syntax, so a guided compose skins them with the guide
 * `callouts.defaultTheme`; guideless composes keep the legacy per-tag map.
 *
 * Anchoring: the first insert anchors to `end` (last live node) with
 * `after`; every later insert chains onto the previously minted id, and
 * follow-up ops (style, list, shading, footnote, endnote) target minted ids.
 * Ids are deterministic (`md_0001`, ...) so repeated compiles of one source
 * give identical op JSON, except image and chart file bytes which depend on
 * file content (ids stay deterministic, fitted image sizes likewise derive
 * from the same file bytes). `composeDocument` appends only, except for one
 * narrow case: a pristine blank template (exactly the two seed placeholder
 * paragraphs and nothing else, see the seed-skip heuristic in this file).
 * There the seeds are removed after the composed chunks commit, so client
 * output holds only composed content. Every other document keeps strict
 * append-only behavior: pre-existing content is never deleted.
 *
 * Template option: `composeDocument` accepts a template name
 * (`report`, `memo`, `blank`). When given, the open document must still hold
 * that template seed text (`Q3 Executive` for report, `MEMORANDUM` for memo,
 * `Blank Document` for blank); otherwise composition fails before any write.
 * The check only gates the write: on a pristine blank template the seeds are
 * still replaced as described above, so a `{ template: 'blank' }` compose
 * proves the seed was present without leaving placeholder text behind.
 *
 * Atomicity: input errors throw from `compileMarkdown` before any batch
 * runs (zero writes). Runtime failures roll back through the engine's own
 * per-batch guarantee. Sources compiling past the per-batch op cap are
 * applied as sequential chunks; chunk N+1 is only attempted after chunk N
 * commits, so a mid-commit runtime failure can leave earlier chunks applied
 * (same exposure as any multi-batch caller; deterministic validation
 * failures still write nothing because they throw at compile time). Inline
 * run splitting runs after all chunks commit; it validates at compile time,
 * and any unexpected runtime failure there restores a pre-batch package
 * snapshot (whole zip including relationships, so no half-split paragraph
 * or orphan relationship survives). Span segments normalize newlines to
 * spaces before runs emit, so `w:t` output never carries raw line breaks:
 * single newlines inside a paragraph read as spaces, matching the plain
 * insert text.
 */

import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  type ExecutionResult,
  ENGINE_MAX_OPS,
  OperationEngine,
} from '../operation-engine';
import { type Operation, validateOperationsStrict } from '../operation-schema';
import type { DocxParser } from '../parser';
import { createWordElement } from '../ooxml/ns';
import { latexToOmml } from '../ooxml/omml';
import { assertSafeLatex, assertValidConvertedParagraph } from '../operations/equation';
import { createTextElement } from '../operations/base';
import { HYPERLINK_REL_TYPE } from '../operations/hyperlink';
import { allocateSafeHyperlinkRId } from '../manipulators/relationships';
import { auditExternalRelTarget } from '../ooxml/xxe-guard';
import { compileMarkdownDetailed, resolveComposeStyleGuide } from './compile';
import { BLANK_SEED_HINT, BLANK_SEED_TITLE, COMPOSE_CODE_FONT, TEMPLATE_SEEDS } from './constants';
import type { StyleGuideConfig } from '../style-guide/types';
import { nlToSpace, segmentsNeedSplit } from './text-spans';
import type { SpanSegment } from './text-spans';

// Barrel: historical deep import stays working; canonical homes above.
export * from './constants';
export * from './text-spans';
export * from './assets';
export * from './compile';

export interface ComposeOptions {
  /** Override the per-batch op cap chunk size (defaults to the engine cap). */
  chunkSize?: number;
  /** Directory that image and chart path references resolve against. Absent = path references fail closed. */
  allowedDir?: string;
  /** Require the open document to still hold this template seed before writing. */
  template?: 'report' | 'memo' | 'blank';
  /**
   * Active style guide: forwarded to the batch engine so composed tables,
   * charts, callouts, and fenced code blocks inherit the house skin
   * (explicit op params still win). Composed Markdown tables take the guide
   * `tables.headerFill` and composed `> [!TAG]` alerts take the guide
   * `callouts.defaultTheme` (Markdown carries no per-table or per-alert
   * theme syntax, so the guide wins there outright). Omitted = the built-in
   * "Default Corporate" guide; `'legacy'` = the old look (gray headers,
   * no header text colour).
   */
  styleGuide?: StyleGuideConfig | 'legacy';
  /** Treat `$...$` as inline math (default off: prices and shell text can look like math). */
  dollarMath?: boolean;
}

/**
 * Pristine blank-template heuristic. Returns the two seed node ids when the
 * open document is still exactly the fresh blank template, otherwise null.
 *
 * The check is deliberately narrow so a real user document can never match:
 * exactly two body nodes, both plain paragraphs, with the exact seed texts
 * (`Blank Document`, `Start typing your content here...`), the seed
 * paragraph styles still `Title`/`Subtitle`, no tracked changes anywhere,
 * and no headers, footers, charts, media parts, footnotes, endnotes, or
 * comments. Any edit (extra node, changed text, restyled seed, tracked
 * change, header/footer, image, note, comment) disables replacement and the
 * compose stays append-only. Report and memo seeds never match (different
 * texts and node counts), so they always keep append-only behavior.
 *
 * Notes guard: a two-node document that carries footnote, endnote, or
 * comment data still matches on text alone, so the guard explicitly requires
 * zero notes. Both the package parts (`word/footnotes.xml`,
 * `word/endnotes.xml`, `word/comments.xml`, `word/commentsExtended.xml`)
 * and in-node reference markup (`w:footnoteReference`,
 * `w:endnoteReference`, `w:commentReference`, `w:commentRangeStart`,
 * `w:commentRangeEnd`) must be absent; either presence keeps append-only so
 * seed deletion can never orphan note parts or leave dangling references.
 *
 * Styles note: the seeds must still carry their template paragraph styles
 * (`Title` on the first, `Subtitle` on the second), checked cheaply from the
 * seed `rawXml`. Extra entries left in `word/styles.xml` by other edits are
 * inert (no seed references them) and do not block replacement; only a
 * restyled seed blocks.
 */
function pristineBlankSeedIds(parser: DocxParser): string[] | null {
  let summary: {
    hasTrackedChanges?: boolean;
    headers?: unknown[];
    footers?: unknown[];
    charts?: unknown[];
    mediaFiles?: Record<string, unknown>;
    nodes?: Array<{
      id?: string;
      type?: string;
      text?: string;
      rawXml?: string;
      hasTrackedChanges?: boolean;
    }>;
  };
  try {
    summary = parser.getStructuralSummary() as typeof summary;
  } catch {
    return null;
  }
  const nodes = summary.nodes ?? [];
  if (nodes.length !== 2) return null;
  if (summary.hasTrackedChanges === true) return null;
  if (Object.keys(summary.mediaFiles ?? {}).length > 0) return null;
  if ((summary.headers ?? []).length > 0) return null;
  if ((summary.footers ?? []).length > 0) return null;
  if ((summary.charts ?? []).length > 0) return null;
  const [first, second] = nodes;
  if (first.type !== 'paragraph' || second.type !== 'paragraph') return null;
  if (first.hasTrackedChanges === true || second.hasTrackedChanges === true) return null;
  if ((first.text ?? '').trim() !== BLANK_SEED_TITLE) return null;
  if ((second.text ?? '').trim() !== BLANK_SEED_HINT) return null;
  const firstRaw = first.rawXml ?? '';
  const secondRaw = second.rawXml ?? '';
  if (firstRaw.length === 0 || secondRaw.length === 0) return null;
  for (const raw of [firstRaw, secondRaw]) {
    if (
      raw.includes('w:footnoteReference') ||
      raw.includes('w:endnoteReference') ||
      raw.includes('w:commentReference') ||
      raw.includes('w:commentRangeStart') ||
      raw.includes('w:commentRangeEnd')
    ) {
      return null;
    }
  }
  if (!firstRaw.includes('w:val="Title"') || !secondRaw.includes('w:val="Subtitle"')) return null;
  try {
    const zip = (
      parser as unknown as { getZipAdapter?: () => { hasFile: (p: string) => boolean } }
    ).getZipAdapter?.();
    if (typeof zip?.hasFile === 'function') {
      for (const part of [
        'word/footnotes.xml',
        'word/endnotes.xml',
        'word/comments.xml',
        'word/commentsExtended.xml',
      ]) {
        try {
          if (zip.hasFile(part)) return null;
        } catch {
          return null;
        }
      }
    } else {
      return null;
    }
  } catch {
    return null;
  }
  const firstId = first.id ?? '';
  const secondId = second.id ?? '';
  if (firstId.length === 0 || secondId.length === 0 || firstId === secondId) return null;
  return [firstId, secondId];
}

function assertTemplateSeed(parser: DocxParser, template: string): void {
  if (template !== 'report' && template !== 'memo' && template !== 'blank') {
    throw new Error(
      'E_VALIDATION: line 1: template must be one of report|memo|blank. Remedy: pass template report, memo, or blank, or omit it.'
    );
  }
  const marker = TEMPLATE_SEEDS[template];
  let texts: string[] = [];
  try {
    const summary = parser.getStructuralSummary() as { nodes?: Array<{ text?: string }> };
    texts = (summary.nodes ?? []).map((n) => n.text ?? '');
  } catch {
    throw new Error(
      'E_VALIDATION: line 1: template check cannot read the open document. Remedy: open a document first, then compose into it.'
    );
  }
  const found = texts.some((t) => t.includes(marker));
  if (!found) {
    throw new Error(
      `E_VALIDATION: line 1: open document does not hold the ${template} seed ("${marker}"). Remedy: open a fresh ${template} document first, then compose into it.`
    );
  }
}

function applyInlineSpans(
  parser: DocxParser,
  ast: { sections: Array<{ nodes: Array<{ id: string; xml?: Element; version: number }> }> } & Record<string, unknown>,
  operations: Operation[]
): void {
  const docParser = parser as unknown as {
    getNode(id: string): Element | undefined;
    getDocumentDom(): Document;
    getZipAdapter(): {
      hasFile(p: string): boolean;
      readText(p: string): string;
      writeText(p: string, text: string): void;
    };
  };
  const liveAst = ast as unknown as {
    relationships: Record<string, { type: string; target: string; targetMode: string }>;
    sections: Array<{ nodes: Array<{ id: string; xml?: Element; version: number }> }>;
  };
  for (const op of operations) {
    const rec = op as unknown as Record<string, unknown>;
    if (rec['action'] !== 'insert' || rec['type'] !== 'paragraph') continue;
    const params = rec['params'] as Record<string, unknown> | undefined;
    const rawSpans = params?.['mdSpans'] as SpanSegment[] | undefined;
    if (!Array.isArray(rawSpans) || !segmentsNeedSplit(rawSpans as SpanSegment[])) continue;
    const spans = rawSpans as SpanSegment[];
    const id = String(rec['id'] ?? '');
    if (!id) continue;
    let node: { id: string; xml?: Element; version: number } | undefined;
    for (const section of liveAst.sections) {
      const found = section.nodes.find((n) => n.id === id);
      if (found) {
        node = found;
        break;
      }
    }
    if (!node?.xml) {
      throw new Error(
        `E_VALIDATION: line 1: span target paragraph '${id}' is missing after compose. Remedy: retry the compose without manual edits in between.`
      );
    }
    const para = node.xml as Element;
    const doc = para.ownerDocument;
    if (!doc) continue;
    const keptFootRefs: Element[] = [];
    const kids = Array.from(para.childNodes);
    for (const kid of kids) {
      if (kid.nodeType !== 1) {
        if (kid.nodeType === 3 && kid.nodeValue && kid.nodeValue.trim().length === 0) {
          para.removeChild(kid);
        }
        continue;
      }
      const el = kid as Element;
      if (el.tagName === 'w:pPr') continue;
      const refs = el.getElementsByTagName('w:footnoteReference');
      const endRefs = el.getElementsByTagName('w:endnoteReference');
      if ((refs && refs.length > 0) || (endRefs && endRefs.length > 0)) {
        keptFootRefs.push(el);
        continue;
      }
      if (el.tagName === 'w:hyperlink') {
        para.removeChild(el);
        continue;
      }
      if (el.tagName === 'w:r' || el.tagName === 'w:ins' || el.tagName === 'w:del') {
        para.removeChild(el);
        continue;
      }
      para.removeChild(el);
    }
    const buildRun = (seg: SpanSegment): Element => {
      const run = createWordElement(doc, 'w:r');
      const needsRPr = seg.bold || seg.italic || seg.code || seg.linkUrl !== undefined;
      if (needsRPr) {
        const rPr = createWordElement(doc, 'w:rPr');
        if (seg.bold) rPr.appendChild(createWordElement(doc, 'w:b'));
        if (seg.italic) rPr.appendChild(createWordElement(doc, 'w:i'));
        if (seg.code) {
          const rFonts = createWordElement(doc, 'w:rFonts');
          rFonts.setAttribute('w:ascii', COMPOSE_CODE_FONT);
          rFonts.setAttribute('w:hAnsi', COMPOSE_CODE_FONT);
          rPr.appendChild(rFonts);
        }
        if (seg.linkUrl !== undefined) {
          const color = createWordElement(doc, 'w:color');
          color.setAttribute('w:val', '0066CC');
          rPr.appendChild(color);
          const underline = createWordElement(doc, 'w:u');
          underline.setAttribute('w:val', 'single');
          rPr.appendChild(underline);
        }
        run.appendChild(rPr);
      }
      run.appendChild(createTextElement(doc, nlToSpace(seg.text)));
      return run;
    };
    const contentEls: Element[] = [];
    for (const seg of spans) {
      if (seg.text.length === 0) continue;
      if (seg.math) {
        try {
          assertSafeLatex(seg.text);
          const math = latexToOmml(doc, seg.text);
          const probe = createWordElement(doc, 'w:p');
          probe.appendChild(math.cloneNode(true));
          assertValidConvertedParagraph(probe);
          contentEls.push(math);
        } catch (err) {
          const raw = err instanceof Error ? err.message : String(err);
          const msg = raw.replace(/^\[?E_VALIDATION\]?:?\s*/, '');
          throw new Error(
            `E_VALIDATION: line ${seg.line ?? 1}: ${msg}${/Remedy:/.test(msg) ? '' : ' Remedy: remove or simplify the math.'}`
          );
        }
        continue;
      }
      if (seg.linkUrl === undefined) {
        contentEls.push(buildRun(seg));
        continue;
      }
      const url = seg.linkUrl;
      // H3 external-rel allowlist (defense in depth behind validateLinkUrl,
      // which already rejects dangerous schemes): the decision is recorded
      // on the xxe-guard audit trail whether allowed or refused.
      const relDecision = auditExternalRelTarget(url);
      if (!relDecision.allowed) {
        throw new Error(
          `E_VALIDATION: line 1: link scheme '${relDecision.scheme}' is not allowed. Remedy: use http, https, mailto, ftp, tel, sms, or a relative address.`
        );
      }
      const relId = allocateSafeHyperlinkRId(liveAst as never);
      liveAst.relationships[relId] = { type: HYPERLINK_REL_TYPE, target: url, targetMode: 'External' };
      try {
        const zipAdapter = docParser.getZipAdapter();
        const relsPath = 'word/_rels/document.xml.rels';
        if (zipAdapter.hasFile(relsPath)) {
          const relsRaw = zipAdapter.readText(relsPath);
          const relsDoc = new DOMParser().parseFromString(relsRaw, 'text/xml');
          const relsEl = relsDoc.getElementsByTagName('Relationships')[0];
          if (relsEl) {
            const rel = relsDoc.createElement('Relationship');
            rel.setAttribute('Id', relId);
            rel.setAttribute('Type', HYPERLINK_REL_TYPE);
            rel.setAttribute('Target', url);
            rel.setAttribute('TargetMode', 'External');
            relsEl.appendChild(rel);
            zipAdapter.writeText(relsPath, new XMLSerializer().serializeToString(relsDoc));
          }
        }
      } catch (err: unknown) {
        throw new Error(
          `E_VALIDATION: line 1: link relationship write failed (${err instanceof Error ? err.message : String(err)}). Remedy: retry the compose.`
        );
      }
      const hyperlink = createWordElement(doc, 'w:hyperlink');
      hyperlink.setAttribute('r:id', relId);
      hyperlink.setAttribute('w:history', '1');
      hyperlink.appendChild(buildRun(seg));
      contentEls.push(hyperlink);
    }
    const refSibs = keptFootRefs;
    for (const el of refSibs) {
      if (el.parentNode === para) para.removeChild(el);
    }
    for (const el of contentEls) para.appendChild(el);
    for (const el of refSibs) para.appendChild(el);
    node.version += 1;
  }
}

/**
 * Compiles Markdown and applies it to an open document through
 * `executeBatch`, chunked at the engine op cap. Appends only, except that a
 * pristine blank template (see the seed-skip heuristic above) has its two
 * seed placeholders removed after the composed chunks commit, so client
 * output holds only composed content. Input errors throw before any write.
 */
export async function composeDocument(
  parser: DocxParser,
  markdown: string,
  options?: ComposeOptions
): Promise<ExecutionResult> {
  if (!parser || typeof parser.getAST !== 'function') {
    throw new Error(
      'E_VALIDATION: line 1: composeDocument needs an open document parser. Remedy: open a document first, then compose into it.'
    );
  }
  if (options?.template !== undefined) {
    assertTemplateSeed(parser, String(options.template));
  }
  const styleGuide = resolveComposeStyleGuide(options?.styleGuide);
  const { operations, warnings: compileWarnings } = compileMarkdownDetailed(markdown, {
    allowedDir: options?.allowedDir,
    ...(options?.styleGuide !== undefined ? { styleGuide: options.styleGuide } : {}),
    ...(options?.dollarMath ? { dollarMath: true } : {}),
  });
  validateOperationsStrict(operations);

  const chunkSize =
    options?.chunkSize !== undefined ? Math.floor(options.chunkSize) : ENGINE_MAX_OPS;
  if (!Number.isInteger(chunkSize) || chunkSize < 1 || chunkSize > ENGINE_MAX_OPS) {
    throw new Error(
      `E_VALIDATION: line 1: chunkSize must be an integer in [1, ${ENGINE_MAX_OPS}]. Remedy: omit chunkSize or pass a value inside that range.`
    );
  }

  const ast = await parser.getAST();
  // Seed-skip capture runs after compile+validate (so input errors still
  // throw with zero writes) but before any batch mutates the document.
  const blankSeedIds = pristineBlankSeedIds(parser);
  // Whole-package snapshot (every zip entry including relationships), so a
  // later rewind leaves no orphan rels or media behind. The narrower
  // document-only snapshot cannot cover hyperlink rels minted during span
  // splitting.
  let preBatchSnapshot: Map<string, Buffer> | null = null;
  try {
    if (typeof parser.capturePackageSnapshot === 'function') {
      preBatchSnapshot = parser.capturePackageSnapshot();
    }
  } catch { preBatchSnapshot = null; }
  const engine = new OperationEngine(ast, parser, styleGuide ? { styleGuide } : undefined);
  const aggregate: ExecutionResult = {
    success: true,
    appliedCount: 0,
    errors: [],
    rollbackTriggered: false,
    skippedIds: [],
    warnings: [...compileWarnings],
    revisionsPending: 0,
    possiblyStaleBodyRefs: [],
  };

  let base = 0;
  for (let start = 0; start < operations.length; start += chunkSize) {
    const chunk = operations.slice(start, start + chunkSize);
    const result = await engine.executeBatch({ operations: chunk, sequential: true });
    aggregate.appliedCount += result.appliedCount;
    for (const entry of result.errors) {
      aggregate.errors.push({
        ...entry,
        operationIndex: entry.operationIndex >= 0 ? base + entry.operationIndex : entry.operationIndex,
      });
    }
    for (const warning of result.warnings) {
      if (!aggregate.warnings.includes(warning)) aggregate.warnings.push(warning);
    }
    for (const id of result.skippedIds) {
      if (!aggregate.skippedIds.includes(id)) aggregate.skippedIds.push(id);
    }
    for (const ref of result.possiblyStaleBodyRefs) {
      if (!aggregate.possiblyStaleBodyRefs.some((entry) => entry.nodeId === ref.nodeId)) {
        aggregate.possiblyStaleBodyRefs.push(ref);
      }
    }
    aggregate.revisionsPending = result.revisionsPending;
    if (result.revisionsRemedy !== undefined) aggregate.revisionsRemedy = result.revisionsRemedy;
    else delete aggregate.revisionsRemedy;
    if (!result.success) {
      aggregate.success = false;
      aggregate.rollbackTriggered = result.rollbackTriggered;
      break;
    }
    base += chunk.length;
  }
  if (!aggregate.success) return aggregate;
  if (blankSeedIds !== null) {
    // Pristine blank template: remove the two seed placeholders now that the
    // composed chunks hold the new content. Deletion runs last so the first
    // insert always has a live `end` anchor, and composed nodes keep the
    // document non-empty throughout. appliedCount intentionally counts only
    // composed ops (deletion is cleanup, not content); warnings, skipped
    // ids, stale refs, and revision counts from the delete merge into the
    // aggregate so delete-side advisories are never dropped.
    const seedCleanup = await engine.executeBatch({
      operations: [{ action: 'delete', target: blankSeedIds } as Operation],
      sequential: true,
    });
    const mergeSeedCleanupMeta = (): void => {
      for (const warning of seedCleanup.warnings ?? []) {
        if (!aggregate.warnings.includes(warning)) aggregate.warnings.push(warning);
      }
      for (const id of seedCleanup.skippedIds ?? []) {
        if (!aggregate.skippedIds.includes(id)) aggregate.skippedIds.push(id);
      }
      for (const ref of seedCleanup.possiblyStaleBodyRefs ?? []) {
        if (!aggregate.possiblyStaleBodyRefs.some((entry) => entry.nodeId === ref.nodeId)) {
          aggregate.possiblyStaleBodyRefs.push(ref);
        }
      }
      aggregate.revisionsPending = seedCleanup.revisionsPending;
      if (seedCleanup.revisionsRemedy !== undefined) aggregate.revisionsRemedy = seedCleanup.revisionsRemedy;
      else delete aggregate.revisionsRemedy;
    };
    let seedsGone = seedCleanup.success;
    if (seedsGone) {
      try {
        const liveIds = new Set(
          (parser.getStructuralSummary().nodes ?? []).map((node) => node.id)
        );
        seedsGone = blankSeedIds.every((id) => !liveIds.has(id));
      } catch {
        seedsGone = false;
      }
    }
    if (!seedsGone) {
      try {
        if (preBatchSnapshot !== null) {
          parser.restorePackageSnapshot(preBatchSnapshot);
        }
      } catch { /* restore best effort */ }
      // Restore rewinds the composed chunks as well, so the count returns
      // to zero: nothing from this call survives. Delete-side revision and
      // advisory metadata still merges so the failure envelope keeps it.
      aggregate.appliedCount = 0;
      mergeSeedCleanupMeta();
      aggregate.success = false;
      aggregate.rollbackTriggered = true;
      for (const entry of seedCleanup.errors) {
        aggregate.errors.push({ ...entry, operationIndex: -1 });
      }
      if (seedCleanup.errors.length === 0) {
        aggregate.errors.push({
          operationIndex: -1,
          error: 'Seed placeholder removal did not take effect.',
          hint: 'Retry the compose; when the blank seeds persist, delete them with a delete batch first, then compose.',
        });
      }
      return aggregate;
    }
    mergeSeedCleanupMeta();
  }
  try {
    applyInlineSpans(parser, ast as never, operations);
  } catch (err: unknown) {
    try {
      if (preBatchSnapshot !== null) {
        parser.restorePackageSnapshot(preBatchSnapshot);
      }
    } catch { /* restore best effort */ }
    const detail = err instanceof Error ? err.message : String(err);
    if (/^E_VALIDATION/.test(detail)) throw err;
    throw new Error(`E_VALIDATION: line 1: inline formatting failed (${detail}). Remedy: simplify bold, italic, code, or link spans and retry.`);
  }
  return aggregate;
}
