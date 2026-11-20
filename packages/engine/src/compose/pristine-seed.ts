import type { DocxParser } from '../parser';

// The pristine-seed check lives in its own module so the cover page operation can use it without
// importing the whole compose pipeline (which imports the operation engine, which imports it).

/**
 * Pristine template placeholder heuristic. Returns the seed node id(s) when the
 * open document is still a fresh pristine template seed, otherwise null.
 *
 * Recognizes:
 * 1. The blank template: one empty paragraph (`<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>`). It is also
 *    recognised after the file was saved and opened again, because a lone empty paragraph holds nothing to lose.
 * 2. Two-node templates:
 *    - memo (`MEMORANDUM`)
 *    - letter (`Letter`)
 *    - proposal (`Proposal`)
 *    - invoice (`INVOICE`)
 * 3. Four-node templates:
 *    - research (`Research Benchmark Report`)
 *    - corporate (`Corporate Strategy Roadmap`)
 *    - bar-chart-3d (`Velocity Audit & Performance`)
 * 4. Five-node templates:
 *    - financial (`Financial Report & Audit`)
 * 5. Six-node templates:
 *    - report (`Q3 Executive Financial Performance Report` + `1. Executive Summary`)
 *    - novel (`Manuscript Title` + `Chapter 1: The Beginning`)
 *    - academic (`Neural Document Architectures: A Comparative Benchmark`)
 *    - newsletter (`The ASTra Dispatch — Monthly Engineering Gazette`)
 *    - api / api-spec (`ASTra Document Engine API Specification`)
 *
 * Strict guards: no tracked changes, no media files, no headers, no footers,
 * no charts, no footnotes, no endnotes, no comments.
 */
export function pristineTemplateSeedIds(parser: DocxParser, requestedTemplate?: string): string[] | null {
  // A package built in this session is recognised by its byte snapshot. A file that was saved and opened again has
  // no snapshot, so only the one case with nothing to lose is accepted without it: a lone empty paragraph.
  const snapshotPristine = parser.hasPristineBuiltinSeed();
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
  if (![1, 2, 4, 5, 6].includes(nodes.length)) return null;
  if (!snapshotPristine && nodes.length !== 1) return null;
  if (summary.hasTrackedChanges === true) return null;
  if (Object.keys(summary.mediaFiles ?? {}).length > 0) return null;
  if ((summary.headers ?? []).length > 0) return null;
  if ((summary.footers ?? []).length > 0) return null;
  if ((summary.charts ?? []).length > 0) return null;

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

  // Guard: no tracked changes or note/comment references in any node
  for (const n of nodes) {
    if (n.hasTrackedChanges === true) return null;
    const raw = n.rawXml ?? '';
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

  const ids = nodes.map((n) => n.id ?? '').filter((id) => id.length > 0);
  if (ids.length !== nodes.length) return null;

  // Case 1: Fresh 1-node empty paragraph placeholder
  if (nodes.length === 1) {
    const node = nodes[0];
    if (node.type !== 'paragraph') return null;
    if ((node.text ?? '').trim() !== '') return null;
    // Without the snapshot the paragraph must be exactly the empty blank seed: no section break, bookmark,
    // field, drawing or run inside it, and no paragraph properties beyond the Normal style.
    if (!snapshotPristine && !/^<w:p(\s[^>]*)?>\s*(<w:pPr>\s*(<w:pStyle w:val="Normal"\s*\/>)?\s*<\/w:pPr>)?\s*<\/w:p>$|^<w:p(\s[^>]*)?\/>$/.test((node.rawXml ?? '').replace(/\sxmlns(:\w+)?="[^"]*"/g, '').trim())) return null;
    return ids;
  }

  // Case 2: 2-node templates (memo, letter, proposal, invoice)
  if (nodes.length === 2) {
    const first = nodes[0];
    const t0 = (first.text ?? '').trim();
    const tmpl = requestedTemplate?.toLowerCase();
    if (t0 === 'MEMORANDUM' && tmpl === 'memo') return ids;
    if (t0 === 'Letter' && tmpl === 'letter') return ids;
    if (t0 === 'Proposal' && tmpl === 'proposal') return ids;
    if (t0 === 'INVOICE' && tmpl === 'invoice') return ids;
    return null;
  }

  // Case 3: 4-node templates (research, corporate, bar-chart-3d)
  if (nodes.length === 4) {
    const t0 = (nodes[0].text ?? '').trim();
    const tmpl = requestedTemplate?.toLowerCase();
    if (t0 === 'Research Benchmark Report' && tmpl === 'research') return ids;
    if (t0 === 'Corporate Strategy Roadmap' && tmpl === 'corporate') return ids;
    if (t0 === 'Velocity Audit & Performance' && tmpl === 'bar-chart-3d') return ids;
    return null;
  }

  // Case 4: 5-node templates (financial)
  if (nodes.length === 5) {
    const t0 = (nodes[0].text ?? '').trim();
    const tmpl = requestedTemplate?.toLowerCase();
    if (t0 === 'Financial Report & Audit' && tmpl === 'financial') return ids;
    return null;
  }

  // Case 5: 6-node templates (report, novel, academic, newsletter, api)
  if (nodes.length === 6) {
    const t0 = (nodes[0].text ?? '').trim();
    const t2 = (nodes[2].text ?? '').trim();
    const tmpl = requestedTemplate?.toLowerCase();
    if (t0 === 'Q3 Executive Financial Performance Report' && t2 === '1. Executive Summary') {
      if (tmpl === 'report') return ids;
      return null;
    }
    if (t0 === 'Manuscript Title' && t2 === 'Chapter 1: The Beginning' && tmpl === 'novel') {
      return ids;
    }
    if (t0 === 'Neural Document Architectures: A Comparative Benchmark' && tmpl === 'academic') {
      return ids;
    }
    if (t0 === 'The ASTra Dispatch — Monthly Engineering Gazette' && tmpl === 'newsletter') return ids;
    if (t0 === 'ASTra Document Engine API Specification' && (tmpl === 'api' || tmpl === 'api-spec')) return ids;
    return null;
  }

  return null;
}

export const pristineBlankSeedIds = pristineTemplateSeedIds;

/** Every template name the check recognises beyond blank, which it identifies by its two seed texts. */
const SEED_TEMPLATE_NAMES = [
  'memo', 'letter', 'proposal', 'invoice', 'research', 'corporate', 'bar-chart-3d', 'financial',
  'report', 'novel', 'academic', 'newsletter', 'api-spec',
] as const;

/**
 * The seed node ids when the document is still a fresh built-in template of any kind, else null.
 * The seed texts identify the template; the name is only needed to tell templates with the same
 * node count apart, so each name is tried in turn.
 */
export function pristineSeedIdsOfAnyTemplate(parser: DocxParser): string[] | null {
  const blank = pristineTemplateSeedIds(parser);
  if (blank !== null) return blank;
  for (const name of SEED_TEMPLATE_NAMES) {
    const ids = pristineTemplateSeedIds(parser, name);
    if (ids !== null) return ids;
  }
  return null;
}
