/**
 * Document census counters (words / paragraphs / tables / images + page estimate).
 *
 * Single source of truth for the CLI `stats` command and the MCP
 * `get_stats` tool: both import these counters, so the two surfaces can
 * never disagree on what a word or a paragraph is. Relocated from
 * `packages/cli/src/cli/io.ts` (which now re-exports them) and extended
 * with table / image tallies for the agent census.
 */

/** Nearest enclosing ancestor with the given tag name, or null. For a given
 *  tag each node has exactly one nearest owner, which is what makes per-owner
 *  counting dedupe (nested text never repeats per ancestor). */
function nearestOwner(el: unknown, tagName: string): unknown {
  let cur = (el as { parentNode?: unknown }).parentNode;
  while (cur && typeof cur === 'object') {
    const e = cur as { tagName?: string; parentNode?: unknown };
    if (e.tagName === tagName) return e;
    cur = e.parentNode;
  }
  return null;
}

/** Stats recipe: body paragraph text plus table cell text (cells joined with spaces).
 *  Headers, footers, footnotes, comments, and field-code instructions are excluded.
 *  Nested tables: every w:t is counted once — a cell's own text is kept separate
 *  from the cells of a table nested inside it, so nested text never repeats per
 *  ancestor. Textbox content (w:txbxContent) counts as its own paragraph run:
 *  a w:t is swept by exactly one nearest w:p owner, so a textbox paragraph's
 *  text is never also grabbed by the host paragraph around it. Paragraphs
 *  counts paragraph nodes only. Pages is a rough 500-words-per-page estimate. */
export function tableCellTextSpaced(node: { metadata?: { text?: string }; xml?: unknown }): string | undefined {
  try {
    const xml = (node as { xml?: { getElementsByTagName?: (n: string) => { length: number; [i: number]: unknown } } }).xml;
    if (!xml || typeof xml.getElementsByTagName !== 'function') return undefined;
    const tcs = xml.getElementsByTagName('w:tc') as unknown as { length: number; [i: number]: Element };
    if (!tcs || tcs.length === 0) return undefined;
    const cells: string[] = [];
    for (let i = 0; i < tcs.length; i++) {
      const tc = tcs[i] as Element;
      const paras = tc.getElementsByTagName('w:p');
      if (paras.length > 0) {
        const parts: string[] = [];
        for (let p = 0; p < paras.length; p++) {
          const para = paras[p] as Element;
          // Nested-table paragraphs belong to a closer w:tc and are counted by
          // that cell's own iteration — skip them here.
          if (nearestOwner(para, 'w:tc') !== tc) continue;
          const runs = para.getElementsByTagName('w:t');
          let t = '';
          for (let r = 0; r < runs.length; r++) {
            const run = runs[r] as Element;
            // A run below a textbox-nested table has a deeper w:tc owner; that
            // cell's own iteration counts it, so keep it out of this text.
            if (nearestOwner(run, 'w:tc') !== tc) continue;
            // Textbox paragraphs own their runs: the host paragraph's sweep
            // skips them here so the textbox w:p's own iteration counts each
            // w:t exactly once instead of twice.
            if (nearestOwner(run, 'w:p') !== para) continue;
            t += run.textContent || '';
          }
          if (t.trim().length > 0) parts.push(t.trim());
        }
        if (parts.length > 0) cells.push(parts.join(' '));
      } else {
        const runs = tc.getElementsByTagName('w:t');
        let t = '';
        for (let r = 0; r < runs.length; r++) {
          const run = runs[r] as Element;
          if (nearestOwner(run, 'w:tc') !== tc) continue;
          t += run.textContent || '';
        }
        if (t.trim().length > 0) cells.push(t.trim());
      }
    }
    if (cells.length === 0) return '';
    return cells.join(' ');
  } catch {
    return undefined;
  }
}

/** Count of nested w:tbl descendants inside a table node's own XML. xmldom's
 *  getElementsByTagName returns descendants only (the node's own w:tbl is
 *  excluded), so a count of 0 means a flat table — flat-document tallies are
 *  unchanged. Non-inspectable XML degrades to 0 instead of throwing. */
function nestedTableCount(node: { xml?: unknown }): number {
  try {
    const xml = (node as { xml?: { getElementsByTagName?: (n: string) => { length: number } } }).xml;
    if (!xml || typeof xml.getElementsByTagName !== 'function') return 0;
    return xml.getElementsByTagName('w:tbl').length;
  } catch {
    return 0;
  }
}

export interface DocStats {
  words: number;
  paragraphs: number;
  pages: number;
  tables: number;
  images: number;
}

/** Stable word/paragraph/table/image/page stats over the AST node texts.
 *  tables counts AST-registered table nodes and the w:tbl elements nested
 *  inside their cells (+1 each); a table that only lives in a textbox/shape
 *  in a plain body paragraph has no AST node and is not counted, so this can
 *  trail a raw document.xml w:tbl census. */
export function computeDocStats(ast: { sections: Array<{ nodes: Array<{ type: string; metadata: { text?: string }; xml?: unknown }> }> }): DocStats {
  let words = 0;
  let paragraphs = 0;
  let tables = 0;
  let images = 0;
  for (const sec of ast.sections) {
    for (const node of sec.nodes) {
      if (node.type === 'paragraph') paragraphs += 1;
      if (node.type === 'table') tables += 1 + nestedTableCount(node);
      if (node.type === 'image') images += 1;
      let text = node.metadata?.text ?? '';
      // Table nodes carry concatenated cell text in metadata.text (no cell
      // boundaries), which fuses edge words. Re-space via w:tc/w:p so the
      // count matches the documented recipe and the markdown export.
      if (node.type === 'table') {
        const spaced = tableCellTextSpaced(node);
        if (spaced !== undefined) text = spaced;
      }
      const hits = text.match(/\S+/g);
      if (hits) words += hits.length;
    }
  }
  // 500 words/page is the standard editorial estimate. The engine page-map is
  // renderer-backed (needs LibreOffice) — stats stays deterministic here.
  return { words, paragraphs, pages: Math.max(1, Math.ceil(words / 500)), tables, images };
}
