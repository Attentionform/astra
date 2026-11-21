import type { PageSetupProperties } from '../ast';
import { PAPERS } from './papers';
import { twipsToInches, twipsToMm } from './length';

/** One length in the two units that people use. */
export interface PageLength {
  twips: number;
  mm: number;
  in: number;
}

/** What a section's page is, as the file has it, for a reader that wants to check what was set. */
export interface PageDescription {
  /** The named paper that the size matches within one twip, or `Custom`. */
  paper: string;
  orientation: 'portrait' | 'landscape';
  width: PageLength;
  height: PageLength;
  margins: Partial<Record<'top' | 'bottom' | 'left' | 'right' | 'header' | 'footer' | 'gutter', PageLength>>;
  columns: number;
  /** One line that says all of it, such as `A4 portrait, 210 x 297 mm (8.27 x 11.69 in); margins top 25.4 mm, ...`. */
  summary: string;
}

const length = (twips: number): PageLength => ({ twips, mm: Math.round(twipsToMm(twips) * 10) / 10, in: twipsToInches(twips) });

/** Describes the page of a section. Returns null when the section has no page size to describe. */
export function describePage(ps: PageSetupProperties | null | undefined): PageDescription | null {
  if (!ps || !ps.pageWidthTwips || !ps.pageHeightTwips) return null;
  const orientation = ps.orientation === 'landscape' || ps.pageWidthTwips > ps.pageHeightTwips ? 'landscape' : 'portrait';
  const margins: PageDescription['margins'] = {};
  const fields: Array<[keyof PageDescription['margins'], number | undefined]> = [
    ['top', ps.marginTop], ['bottom', ps.marginBottom], ['left', ps.marginLeft], ['right', ps.marginRight],
    ['header', ps.headerDistance], ['footer', ps.footerDistance], ['gutter', ps.gutter],
  ];
  for (const [name, value] of fields) if (value !== undefined) margins[name] = length(value);
  const width = length(ps.pageWidthTwips);
  const height = length(ps.pageHeightTwips);
  const paper = ps.paperSize ?? 'Custom';
  const m = (name: keyof PageDescription['margins']): string => (margins[name] ? `${name} ${margins[name]!.mm} mm` : '');
  const list = (['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'] as const).map(m).filter((s) => s !== '');
  const columns = ps.columns ?? 1;
  const summary = `${paper} ${orientation}, ${width.mm} x ${height.mm} mm (${width.in} x ${height.in} in); margins ${list.join(', ')}${columns > 1 ? `; ${columns} columns` : ''}`;
  return { paper, orientation, width, height, margins, columns, summary };
}

/** The names of the papers, for a caller that lists them. */
export const PAPER_LIST = PAPERS.map((p) => ({ name: p.name, label: p.label }));
