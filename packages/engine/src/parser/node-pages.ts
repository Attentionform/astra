import { linesPerPageFromSectPr } from './sect-pr';

/**
 * IN-07: renderer-compatible embedded nodePages for getStructuralSummary.
 * Explicit page-break nodes consume exactly like the canvas paginator
 * (break maps to the new page); density fallback uses geometry-scaled
 * lines-per-page with the same ~12 words/line + table/image weights as
 * visual_review so the two estimates cannot diverge from each other.
 * Exported for unit tests.
 */
export function buildEmbeddedNodePages(
  nodes: { id: string; type?: string; preview?: string; text?: string; metadata?: Record<string, unknown>; rawXml?: string }[],
  sectPrXml?: string
): { nodePages: Record<string, number>; totalPages: number; linesPerPage: number } {
  const nodePages: Record<string, number> = {};
  let currentPage = 1;
  let lineCount = 0;
  const LINES_PER_PAGE = linesPerPageFromSectPr(sectPrXml);
  const list = Array.isArray(nodes) ? nodes : [];
  for (const node of list) {
    const rec = node as unknown as Record<string, unknown>;
    const id = rec['id'];
    if (typeof id !== 'string' || id.length === 0) continue;
    const raw: string = typeof rec['rawXml'] === 'string' ? (rec['rawXml'] as string) : '';
    const nodeType = rec['type'];
    const preview = rec['preview'];
    const isBreak =
      nodeType === 'pageBreak' ||
      preview === '[Page Break]' ||
      (typeof raw === 'string' && raw.includes('w:type="page"'));
    if (isBreak) {
      currentPage++;
      lineCount = 0;
      nodePages[id] = currentPage;
      continue;
    }
    nodePages[id] = currentPage;
    const textVal = rec['text'];
    const previewVal = rec['preview'];
    const text: string = typeof textVal === 'string' ? textVal : typeof previewVal === 'string' ? previewVal : '';
    if (nodeType === 'table') {
      const meta = rec['metadata'] as { tableDimensions?: { rows?: unknown } } | undefined;
      const rowsVal = meta?.tableDimensions?.rows;
      const rows = typeof rowsVal === 'number' && Number.isFinite(rowsVal) ? rowsVal : 3;
      lineCount += rows + 2;
    } else if (nodeType === 'image' || nodeType === 'chart3d' || nodeType === 'chart') {
      lineCount += 15;
    } else {
      const words = String(text).split(/\s+/).filter(Boolean).length;
      lineCount += Math.max(1, Math.ceil(words / 12));
    }
    if (lineCount >= LINES_PER_PAGE) {
      currentPage++;
      lineCount = 0;
    }
  }
  return { nodePages, totalPages: currentPage, linesPerPage: LINES_PER_PAGE };
}
