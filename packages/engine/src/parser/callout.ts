import { findDirectChild } from './dom-utils';

/**
 * Persisted-callout recognition for `getAST()`.
 *
 * The callout op (`operations/callout.ts`) builds a single-cell table with
 * a thick colored left border, a tinted cell background, and a bold colored
 * title run. The save path strips every `data-astra-id` (`readDocumentXml`),
 * so the `callout_*` id never reaches disk — the structure below is the
 * only signal that survives save/reopen. Matching it here rebuilds the
 * `metadata.callout` marker live sessions carry, so reopened callouts
 * export exactly like live ones (div, not a data table).
 *
 * Fail-safe direction is "plain table": anything edited past recognition
 * (extra rows/cells/paragraphs, changed borders or shading) renders as a
 * table, never a wrong div. Colors are keyed as an exact triple
 * (border + shading + title run must agree on one theme), so an ordinary
 * data table cannot trip the detector by accident.
 */
interface PersistedCallout {
  theme: string;
  title: string | undefined;
  /** Live-canonical text form (`Title: body`), so exporters need no branch. */
  text: string;
}

const THEME_BY_BORDER: Record<string, { theme: string; bg: string; titleColor: string }> = {
  '4F46E5': { theme: 'indigo', bg: 'EEF2FF', titleColor: '3730A3' },
  '059669': { theme: 'emerald', bg: 'ECFDF5', titleColor: '065F46' },
  'D97706': { theme: 'amber', bg: 'FFFBEB', titleColor: '92400E' },
  '475569': { theme: 'slate', bg: 'F8FAFC', titleColor: '1E293B' },
  'E11D48': { theme: 'rose', bg: 'FFF1F2', titleColor: '9F1239' },
};

function directChildren(parent: Element, tag: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i] as Element;
    if (child.nodeType === 1 && child.tagName === tag) out.push(child);
  }
  return out;
}

function paraText(para: Element): string {
  const runs = para.getElementsByTagName('w:t');
  let text = '';
  for (let i = 0; i < runs.length; i++) text += runs[i].textContent || '';
  return text;
}

export function detectPersistedCallout(tbl: Element): PersistedCallout | undefined {
  if (!tbl || tbl.tagName !== 'w:tbl') return undefined;

  // Single row, single cell.
  const rows = directChildren(tbl, 'w:tr');
  if (rows.length !== 1) return undefined;
  const cells = directChildren(rows[0], 'w:tc');
  if (cells.length !== 1) return undefined;
  const tc = cells[0];

  // Accent-bar borders: left single/32pt in a known theme color, rest none.
  const tblPr = findDirectChild(tbl, 'w:tblPr');
  if (!tblPr) return undefined;
  const borders = findDirectChild(tblPr, 'w:tblBorders');
  if (!borders) return undefined;
  const left = findDirectChild(borders, 'w:left');
  if (!left || left.getAttribute('w:val') !== 'single' || left.getAttribute('w:sz') !== '32') return undefined;
  const theme = THEME_BY_BORDER[(left.getAttribute('w:color') || '').toUpperCase()];
  if (!theme) return undefined;
  for (const edge of ['w:top', 'w:bottom', 'w:right']) {
    const el = findDirectChild(borders, edge);
    if (!el || el.getAttribute('w:val') !== 'none') return undefined;
  }

  // Cell shading must agree with the border theme.
  const tcPr = findDirectChild(tc, 'w:tcPr');
  const shd = tcPr ? findDirectChild(tcPr, 'w:shd') : null;
  if (!shd || (shd.getAttribute('w:fill') || '').toUpperCase() !== theme.bg) return undefined;

  // Cell holds only paragraphs: one (untitled body) or two (title + body).
  const cellKids: Element[] = [];
  for (let i = 0; i < tc.childNodes.length; i++) {
    const child = tc.childNodes[i] as Element;
    if (child.nodeType !== 1) continue;
    if (child.tagName !== 'w:p' && child.tagName !== 'w:tcPr') return undefined;
    if (child.tagName === 'w:p') cellKids.push(child);
  }
  if (cellKids.length < 1 || cellKids.length > 2) return undefined;

  let title: string | undefined;
  let body: string;
  if (cellKids.length === 2) {
    // Title paragraph: bold run in the theme title color.
    const first = cellKids[0];
    if (first.getElementsByTagName('w:b').length === 0) return undefined;
    const colors = first.getElementsByTagName('w:color');
    let titleColored = false;
    for (let i = 0; i < colors.length; i++) {
      if (((colors[i] as unknown as Element).getAttribute('w:val') || '').toUpperCase() === theme.titleColor) {
        titleColored = true;
        break;
      }
    }
    if (!titleColored) return undefined;
    title = paraText(first).trim() || undefined;
    body = paraText(cellKids[1]);
  } else {
    body = paraText(cellKids[0]);
  }
  body = body.trim();
  if (!title && !body) return undefined;

  return { theme: theme.theme, title, text: title ? `${title}: ${body}` : body };
}
