import type { DocxParser } from '../parser';

// The blank-document check lives in its own module so the cover page operation can use it without importing the whole
// compose pipeline (which imports the operation engine, which imports it). Designed templates are recognised by their
// marker and manifest instead (template-seed.ts).

const LONE_EMPTY_PARAGRAPH = /^<w:p(\s[^>]*)?>\s*(<w:pPr>\s*(<w:pStyle w:val="Normal"\s*\/>)?\s*<\/w:pPr>)?\s*<\/w:p>$|^<w:p(\s[^>]*)?\/>$/;

/**
 * The id of the only body node when the open document is one empty paragraph, otherwise null.
 *
 * That is what a blank document is. It holds nothing to lose, so compose and a cover page may replace it, also after the file
 * was saved and opened again. The paragraph must be exactly empty: no run, no section break, bookmark, field or drawing, and no
 * paragraph properties beyond the Normal style. The document must also hold no tracked change, header, footer, chart, media
 * file, footnote, endnote or comment.
 */
export function loneEmptyParagraphIds(parser: DocxParser): string[] | null {
  let summary: {
    hasTrackedChanges?: boolean;
    headers?: unknown[];
    footers?: unknown[];
    charts?: unknown[];
    mediaFiles?: Record<string, unknown>;
    nodes?: Array<{ id?: string; type?: string; text?: string; rawXml?: string; hasTrackedChanges?: boolean }>;
  };
  try {
    summary = parser.getStructuralSummary() as typeof summary;
  } catch {
    return null;
  }
  const nodes = summary.nodes ?? [];
  if (nodes.length !== 1) return null;
  if (summary.hasTrackedChanges === true) return null;
  if (Object.keys(summary.mediaFiles ?? {}).length > 0) return null;
  if ((summary.headers ?? []).length > 0) return null;
  if ((summary.footers ?? []).length > 0) return null;
  if ((summary.charts ?? []).length > 0) return null;

  try {
    const zip = parser.getZipAdapter();
    for (const part of ['word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml', 'word/commentsExtended.xml']) {
      if (zip.hasFile(part)) return null;
    }
  } catch {
    return null;
  }

  const node = nodes[0];
  if (!node.id || node.type !== 'paragraph' || node.hasTrackedChanges === true) return null;
  if ((node.text ?? '').trim() !== '') return null;
  if (!LONE_EMPTY_PARAGRAPH.test((node.rawXml ?? '').replace(/\sxmlns(:\w+)?="[^"]*"/g, '').trim())) return null;
  return [node.id];
}
