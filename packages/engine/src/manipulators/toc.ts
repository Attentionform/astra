import { z } from 'zod';
import * as crypto from 'crypto';
import { createWordElement } from '../ooxml/ns';
import { SETTINGS_CHILD_ORDER, appendInSchemaOrder } from '../ooxml/schema-order';
import { DocumentAST, Node, calculateNodeChecksum } from '../ast';
import { DocxParser } from '../parser';
import { buildPageMap } from '../page-map';
import { allocateNextBookmarkId } from './relationships';
import { sanitizeXml } from '../validator';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { resolveReproducibleMode, resolveDeterministicZipDate } from '../zip-hardening';

/** Deterministic by default; wall-clock only when opted out. */
function isTocReproducible(): boolean {
  try {
    return resolveReproducibleMode(undefined);
  } catch {
    return true;
  }
}

/** Content-seeded 8-hex for TOC container ids (stable across runs). */
function tocHash8(input: string): string {
  try {
    return crypto.createHash('sha256').update(input).digest('hex').slice(0, 8);
  } catch {
    return '00000000';
  }
}

export interface TOCEntry {
  headingText: string;
  level: number;
  bookmarkName: string;
  pageNumber: number;
  nodeId: string;
  /** G2-7: true when pageNumber came from the heuristic estimate fallback. */
  pageEstimate?: boolean;
  /** G2-7: provenance of the page number ('renderer' authoritative vs 'estimate'). */
  pageSource?: 'renderer' | 'estimate';
  /** Sprint-3 footnotes path: true when the heading hosts a footnote anchor. */
  hasFootnoteAnchor?: boolean;
}

/**
 * Sprint-3 strict TOC entry schema. Rejects unknown keys, empty headings,
 * out-of-range levels/pages, and over-long strings at the gate (with remedy)
 * instead of emitting corrupt TOC XML.
 */
export const TOCEntrySchema = z
  .object({
    headingText: z.string().min(1).max(2000),
    level: z.number().int().min(1).max(9),
    bookmarkName: z.string().min(1).max(256),
    pageNumber: z.number().int().positive().max(100000),
    nodeId: z.string().min(1).max(256),
    pageEstimate: z.boolean().optional(),
    pageSource: z.enum(['renderer', 'estimate']).optional(),
    hasFootnoteAnchor: z.boolean().optional(),
  })
  .strict();

export type TOCEntryInput = z.infer<typeof TOCEntrySchema>;

function formatTocZodIssues(
  issues: ReadonlyArray<{ path: ReadonlyArray<PropertyKey>; message: string; code: string }>
): string {
  return issues.map((i) => `${i.path.map(String).join('.') || '(root)'}: ${i.message} [${i.code}]`).join('; ');
}

/** Strict TOC entry-list validator. Returns ok:true with parsed entries or throws with remedy. */
export function validateTOCEntriesStrict(entries: unknown): { ok: true; entries: TOCEntry[] } {
  if (!Array.isArray(entries)) {
    throw new Error(
      `E_VALIDATION: TOC entries must be an array (got ${typeof entries}). Remedy: pass TOCEntry[] from TOCManipulator.scanHeadings().`
    );
  }
  const parsed: TOCEntry[] = [];
  const failures: string[] = [];
  for (let i = 0; i < entries.length; i++) {
    const r = TOCEntrySchema.safeParse(entries[i]);
    if (r.success) parsed.push(r.data as TOCEntry);
    else failures.push(`entries[${i}]: ${formatTocZodIssues(r.error.issues)}`);
  }
  if (failures.length > 0) {
    throw new Error(
      `E_VALIDATION: ${failures.length} invalid TOC entr${failures.length === 1 ? 'y' : 'ies'}. ${failures.slice(0, 3).join(' | ')}${failures.length > 3 ? ` | …+${failures.length - 3} more` : ''}. ` +
        `Remedy: fix headingText/level/bookmarkName/pageNumber/nodeId to the TOCEntry schema and retry.`
    );
  }
  return { ok: true, entries: parsed };
}

/**
 * CT_SimpleField siblings guard (TOC-01): a <w:fldSimple> may only contain
 * run-level content (w:r / w:hyperlink) — never block-level <w:p>. TOC
 * entries are emitted as SIBLING w:p elements below the container, never
 * nested inside the field. Fails closed with a remedy when violated.
 */
export function assertFldSimpleHasNoParagraphs(containerXml: string): { ok: true } {
  sanitizeXml(containerXml);
  const doc = new DOMParser({
    errorHandler: {
      warning: (w) => {
        throw new Error(String(w));
      },
      error: (e) => {
        throw new Error(String(e));
      },
      fatalError: (f) => {
        throw new Error(String(f));
      },
    },
  }).parseFromString(`<root>${containerXml}</root>`, 'text/xml');
  const parseErrors = doc.getElementsByTagName('parsererror');
  if (parseErrors.length > 0) {
    throw new Error(
      `E_VALIDATION: TOC container XML is malformed (${parseErrors[0]?.textContent || 'parsererror'}). Remedy: emit via TOCManipulator.generateTOCNode and retry.`
    );
  }
  const fields = doc.getElementsByTagName('w:fldSimple');
  for (let i = 0; i < fields.length; i++) {
    const nested = (fields[i] as Element).getElementsByTagName('w:p');
    if (nested.length > 0) {
      throw new Error(
        `E_VALIDATION: TOC-01 violation — <w:p> nested inside <w:fldSimple> (field ${i}). Remedy: emit TOC entries as sibling w:p elements below the container (see TOCManipulator.generateTOCNode), never as children of fldSimple.`
      );
    }
  }
  return { ok: true };
}

/** Full container check: XXE/script gate + well-formedness + TOC-01 siblings rule. */
export function validateTocContainerXml(containerXml: string): { ok: true } {
  return assertFldSimpleHasNoParagraphs(containerXml);
}

/**
 * Footnotes path: heading text excluding footnote/endnote reference runs.
 * Reference runs carry no <w:t>, but headings edited around a footnote anchor
 * must never leak the anchor id or the footnote body into the TOC label —
 * so text is collected only from runs WITHOUT a footnote/endnote reference
 * child. Falls back to `fallback` when nothing remains.
 */
export function extractHeadingTextExcludingFootnoteRefs(
  nodeXml: Element | null | undefined,
  fallback = ''
): string {
  try {
    if (!nodeXml) return fallback;
    const runs = nodeXml.getElementsByTagName('w:r');
    let text = '';
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i] as Element;
      if (
        run.getElementsByTagName('w:footnoteReference').length > 0 ||
        run.getElementsByTagName('w:endnoteReference').length > 0
      ) {
        continue;
      }
      const ts = run.getElementsByTagName('w:t');
      for (let j = 0; j < ts.length; j++) text += ts[j].textContent || '';
    }
    const trimmed = text.trim();
    return trimmed || fallback;
  } catch {
    return fallback;
  }
}

export class TOCManipulator {
  /**
   * Scans all headings in the AST, assigns/ensures bookmark anchors,
   * and builds structured TOC entry definitions.
   *
   * G2-7 unlabeled-pages fix: page numbers are ESTIMATES unless a live
   * renderer nodePages map is supplied (canvas paginator PageFrames / DOM
   * data-page-number, or parser.getStructuralSummary().nodePages which is
   * geometry-aware but still an estimate). The heuristic word/line fallback
   * may place headings on a neighboring page — every entry therefore carries
   * pageSource/pageEstimate, and generateTOCNode embeds an XML comment +
   * generator note so the estimate is never mistaken for a renderer layout.
   * Word recalculates exact numbers on open via updateFields (see
   * enableWordAutoUpdateFields); prefer the renderer map whenever it is cheap
   * to obtain (field.ts passes parser.getStructuralSummary().nodePages).
   */
  public static scanHeadings(
    ast: DocumentAST,
    maxLevels = 3,
    rendererNodePages?: Record<string, number> | null
  ): TOCEntry[] {
    const pageMap = buildPageMap(ast, rendererNodePages ?? undefined);
    const pageSource: 'renderer' | 'estimate' = pageMap.isEstimate ? 'estimate' : 'renderer';
    const entries: TOCEntry[] = [];
    let tocCounter = 1;

    for (const section of ast.sections) {
      for (const node of section.nodes) {
        if (node.type !== 'paragraph' && !node.id.startsWith('p_')) continue;

        let headingLevel = 0;
        const style = (node.metadata?.namedStyle || '').toLowerCase();
        if (style.startsWith('heading') || style.startsWith('h1') || style.startsWith('h2') || style.startsWith('h3')) {
          const match = style.match(/\d+/);
          headingLevel = match ? parseInt(match[0], 10) : 1;
        }

        if (headingLevel > 0 && headingLevel <= maxLevels) {
          // Footnotes path: metadata.text is w:t-only (reference runs carry
          // no text), but fall back to the footnote-aware extractor so a
          // heading whose only runs host footnote anchors still labels
          // correctly without leaking anchor ids or footnote bodies.
          const text = (node.metadata?.text || '').trim()
            || extractHeadingTextExcludingFootnoteRefs(node.xml, '').trim();
          if (!text) continue;
          const hasFootnoteAnchor = node.xml
            ? node.xml.getElementsByTagName('w:footnoteReference').length > 0
            : false;

          // Generate or retrieve TOC bookmark name. Deterministic default:
          // content-seeded _Toc<hash>_<counter> (stable across runs);
          // wall-clock only when opted out.
          let bookmarkName = isTocReproducible()
            ? `_Toc${tocHash8(`${text}:${headingLevel}:${node.id}`)}_${tocCounter}`
            : `_Toc${Date.now()}_${tocCounter}`;
          if (node.xml) {
            const existingBm = node.xml.getElementsByTagName('w:bookmarkStart')[0];
            if (existingBm) {
              const name = existingBm.getAttribute('w:name');
              if (name && name.startsWith('_Toc')) {
                bookmarkName = name;
              }
            } else {
              // Stamp bookmark into heading XML. OX-12/BX-05: allocate the w:id
              // via the shared safe allocator (max in-use + 1) instead of a
              // fixed 100+counter offset — repeated scans previously re-minted
              // identical ids, producing duplicate bookmarks.
              const doc = node.xml.ownerDocument;
              if (doc) {
                const freshId = String(allocateNextBookmarkId(doc as unknown as Document));
                const bmStart = createWordElement(doc, 'w:bookmarkStart');
                bmStart.setAttribute('w:id', freshId);
                bmStart.setAttribute('w:name', bookmarkName);

                const bmEnd = createWordElement(doc, 'w:bookmarkEnd');
                bmEnd.setAttribute('w:id', freshId);

                if (node.xml.firstChild) {
                  // Insert bookmarkStart after pPr when present (pPr must remain first child per CT_P)
                  const pPrEl = node.xml.getElementsByTagName('w:pPr')[0];
                  const directPPr = pPrEl && pPrEl.parentNode === node.xml ? pPrEl : null;
                  if (directPPr) {
                    node.xml.insertBefore(bmStart, directPPr.nextSibling);
                  } else {
                    node.xml.insertBefore(bmStart, node.xml.firstChild);
                  }
                } else {
                  node.xml.appendChild(bmStart);
                }
                node.xml.appendChild(bmEnd);
              }
            }
          }

          const pageNumber = pageMap.nodePages[node.id] || 1;
          entries.push({
            headingText: text,
            level: headingLevel,
            bookmarkName,
            pageNumber,
            nodeId: node.id,
            pageEstimate: pageMap.isEstimate,
            pageSource,
            ...(hasFootnoteAnchor ? { hasFootnoteAnchor: true as const } : {}),
          });
          tocCounter++;
        }
      }
    }

    return entries;
  }

  /**
   * Generates a fully compliant OOXML Table of Contents container paragraph
   * with dot leaders, tab stops, hyperlinks, and page numbers.
   *
   * G2-7: baked page numbers are labeled estimates — an XML comment is stamped
   * as the first child of the container plus a generator note in metadata
   * (tocPageSource/tocPageEstimate), and Word is instructed to recalculate
   * exact numbers on open via updateFields. Callers should pass the renderer
   * map into scanHeadings when cheap so entries carry pageSource 'renderer'.
   */
  public static async generateTOCNode(
    doc: Document,
    title = 'Table of Contents',
    entries: TOCEntry[],
    fieldId?: string,
    levels = 3
  ): Promise<Node[]> {
    // Sprint-3 strict gate: malformed entries fail closed here with a remedy
    // instead of emitting corrupt TOC XML downstream.
    if (entries.length > 0) validateTOCEntriesStrict(entries);
    const id = fieldId || (isTocReproducible()
      ? `toc_${resolveDeterministicZipDate().getTime()}_${tocHash8(`${title}:${entries.length}:${entries.map((e) => `${e.headingText}:${e.level}:${e.nodeId}`).join('|')}`).slice(0, 6)}`
      : `toc_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`);
    const serializer = new XMLSerializer();

    const tocContainer = createWordElement(doc, 'w:p');
    tocContainer.setAttribute('data-astra-id', id);

    // G2-7: label baked page numbers as estimates in the XML itself (first
    // child comment) so the docx never masquerades a heuristic as a renderer
    // layout. Word recalculates exact numbers via updateFields on open.
    try {
      const estimateNote = doc.createComment(
        ' ASTra TOC: baked page numbers are ESTIMATES (heuristic fallback unless scanHeadings received a renderer nodePages map); Word updates fields on open (updateFields=true). '
      );
      tocContainer.appendChild(estimateNote);
    } catch { /* comments are best-effort — container stays valid without one */ }

    // Title / TOC Heading style
    const pPr = createWordElement(doc, 'w:pPr');
    const pStyle = createWordElement(doc, 'w:pStyle');
    pStyle.setAttribute('w:val', 'TOCHeading');
    pPr.appendChild(pStyle);
    tocContainer.appendChild(pPr);

    const titleR = createWordElement(doc, 'w:r');
    const titleRPr = createWordElement(doc, 'w:rPr');
    titleRPr.appendChild(createWordElement(doc, 'w:b'));
    const szEl = createWordElement(doc, 'w:sz');
    szEl.setAttribute('w:val', '32'); // 16pt
    titleRPr.appendChild(szEl);
    titleR.appendChild(titleRPr);

    const titleT = createWordElement(doc, 'w:t');
    titleT.textContent = title;
    titleR.appendChild(titleT);
    tocContainer.appendChild(titleR);

    // Simple Field Code: TOC \o "1-3" \h \z \u
    // CT_SimpleField may only contain run-level content (w:r / w:hyperlink),
    // never block-level w:p. Keep a single cached-result run here and emit
    // each TOC entry as a sibling w:p below.
    const fldSimple = createWordElement(doc, 'w:fldSimple');
    const topLevel = Number.isInteger(levels) && levels >= 1 && levels <= 9 ? levels : 3;
    fldSimple.setAttribute('w:instr', `TOC \\o "1-${topLevel}" \\h \\z \\u`);
    const cacheR = createWordElement(doc, 'w:r');
    const cacheT = createWordElement(doc, 'w:t');
    cacheT.setAttribute('xml:space', 'preserve');
    cacheT.textContent = '';
    cacheR.appendChild(cacheT);
    fldSimple.appendChild(cacheR);

    tocContainer.appendChild(fldSimple);

    // TOC-01 self-check (fail-closed): the container's fldSimple must never
    // carry a nested w:p — entries are siblings below, asserted here so a
    // future edit cannot silently re-nest them.
    try {
      assertFldSimpleHasNoParagraphs(serializer.serializeToString(tocContainer));
    } catch (err) {
      throw new Error(
        `E_VALIDATION: generated TOC container violates TOC-01 (${err instanceof Error ? err.message : String(err)})`
      );
    }

    const xmlString = serializer.serializeToString(tocContainer);
    const checksum = await calculateNodeChecksum(xmlString);

    // Generator note: provenance of the baked numbers (renderer vs estimate).
    const tocPageSource: 'renderer' | 'estimate' = entries.length > 0 && entries.every((e) => (e as TOCEntry).pageSource === 'renderer')
      ? 'renderer'
      : 'estimate';
    const tocNode: Node = {
      id,
      type: 'paragraph',
      xml: tocContainer,
      metadata: {
        text: title,
        hasCitation: false,
        hasComment: false,
        hasBookmark: false,
        hasTrackedChanges: false,
        hasEquation: false,
        hasField: true,
        hasImage: false,
        hasHyperlink: false,
        namedStyle: 'TOCHeading',
        // G2-7 generator note (minimum labeling when XML comments are stripped):
        tocPageSource,
        tocPageEstimate: tocPageSource === 'estimate',
        tocPageNote: tocPageSource === 'estimate'
          ? 'Baked TOC page numbers are estimates; Word recalculates on open (updateFields).'
          : 'Baked TOC page numbers from the renderer-compatible layout.',
      },
      version: 1,
      checksum
    };

    const tocNodes: Node[] = [tocNode];

    // Default entries if document has no headings yet (placeholders are estimates).
    const displayEntries: TOCEntry[] = entries.length > 0 ? entries : [
      { headingText: '1. Executive Summary', level: 1, bookmarkName: '_Toc1', pageNumber: 1, nodeId: 'p_1', pageEstimate: true, pageSource: 'estimate' },
      { headingText: '2. Key Findings & Clinical Impact', level: 1, bookmarkName: '_Toc2', pageNumber: 2, nodeId: 'p_2', pageEstimate: true, pageSource: 'estimate' },
      { headingText: '3. Evidence-Based Interventions', level: 1, bookmarkName: '_Toc3', pageNumber: 3, nodeId: 'p_3', pageEstimate: true, pageSource: 'estimate' }
    ];

    let entryIdx = 1;
    for (const entry of displayEntries) {
      const itemP = createWordElement(doc, 'w:p');
      const entryId = `${id}_entry_${entryIdx++}`;
      itemP.setAttribute('data-astra-id', entryId);
      const itemPPr = createWordElement(doc, 'w:pPr');

      // Style: TOC1, TOC2, TOC3
      const itemStyle = createWordElement(doc, 'w:pStyle');
      itemStyle.setAttribute('w:val', `TOC${entry.level}`);
      itemPPr.appendChild(itemStyle);

      // Indentation for hierarchical subheadings
      if (entry.level > 1) {
        const ind = createWordElement(doc, 'w:ind');
        ind.setAttribute('w:left', String((entry.level - 1) * 280));
        itemPPr.appendChild(ind);
      }

      // Right-aligned tab stop with dot leader at position 9360 DXA (6.5 inches)
      const tabs = createWordElement(doc, 'w:tabs');
      const tab = createWordElement(doc, 'w:tab');
      tab.setAttribute('w:val', 'right');
      tab.setAttribute('w:leader', 'dot');
      tab.setAttribute('w:pos', '9360');
      tabs.appendChild(tab);
      itemPPr.appendChild(tabs);

      itemP.appendChild(itemPPr);

      // Hyperlink wrapping title + tab leader + page number
      const hyperlink = createWordElement(doc, 'w:hyperlink');
      hyperlink.setAttribute('w:anchor', entry.bookmarkName);
      hyperlink.setAttribute('w:history', '1');

      // Heading Text Run
      const textR = createWordElement(doc, 'w:r');
      if (entry.level === 1) {
        const textRPr = createWordElement(doc, 'w:rPr');
        textRPr.appendChild(createWordElement(doc, 'w:b'));
        textR.appendChild(textRPr);
      }
      const textT = createWordElement(doc, 'w:t');
      textT.textContent = entry.headingText;
      textR.appendChild(textT);
      hyperlink.appendChild(textR);

      // Dot Leader Tab Run
      const tabR = createWordElement(doc, 'w:r');
      tabR.appendChild(createWordElement(doc, 'w:tab'));
      hyperlink.appendChild(tabR);

      // Page Number Run
      const pageR = createWordElement(doc, 'w:r');
      const pageT = createWordElement(doc, 'w:t');
      pageT.textContent = String(entry.pageNumber);
      pageR.appendChild(pageT);
      hyperlink.appendChild(pageR);

      itemP.appendChild(hyperlink);

      const entryXmlString = serializer.serializeToString(itemP);
      const entryChecksum = await calculateNodeChecksum(entryXmlString);

      tocNodes.push({
        id: entryId,
        type: 'paragraph',
        xml: itemP,
        metadata: {
          text: `${entry.headingText} ............ ${entry.pageNumber}`,
          hasCitation: false,
          hasComment: false,
          hasBookmark: false,
          hasTrackedChanges: false,
          hasEquation: false,
          hasField: false,
          hasImage: false,
          hasHyperlink: true,
          namedStyle: `TOC${entry.level}`,
          tocPageSource: (entry as TOCEntry).pageSource ?? 'estimate',
          tocPageEstimate: (entry as TOCEntry).pageEstimate ?? true,
        },
        version: 1,
        checksum: entryChecksum
      });
    }

    return tocNodes;
  }

  /**
   * Sets <w:updateFields w:val="true"/> in settings.xml so Microsoft Word
   * automatically prompts/updates all TOC and page fields when the user opens the file.
   */
  public static enableWordAutoUpdateFields(parser?: DocxParser): void {
    if (!parser) return;
    try {
      const zip = parser.getZipAdapter();
      let settingsXml = zip.readTextOrNull('word/settings.xml');
      if (!settingsXml) {
        settingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:updateFields w:val="true"/>
</w:settings>`;
        zip.writeText('word/settings.xml', settingsXml);
      } else if (!settingsXml.includes('w:updateFields')) {
        // Schema order matters: updateFields sits before footnotePr, endnotePr and compat,
        // so it cannot simply be appended once the part carries those.
        const doc = new DOMParser().parseFromString(settingsXml, 'text/xml');
        const root = doc.documentElement as unknown as Element;
        const updateFields = createWordElement(doc, 'w:updateFields');
        updateFields.setAttribute('w:val', 'true');
        appendInSchemaOrder(root, updateFields, SETTINGS_CHILD_ORDER);
        zip.writeText('word/settings.xml', new XMLSerializer().serializeToString(doc));
      }
    } catch {
      // Non-fatal if settings.xml is unavailable
    }
  }
}
