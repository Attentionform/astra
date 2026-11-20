/**
 * The named papers. One table serves the operation, document creation, the read-back and the docs.
 *
 * Twips are rounded from the size in millimetres or inches. Word's own table can differ by one twip for some papers, so a
 * page within one twip of a named paper is still called by that name.
 */

import { MIN_TEXT_HEIGHT, MIN_TEXT_WIDTH } from './limits';

/** Margins and header and footer distances, in twips. */
export interface PageMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
}

export interface Paper {
  /** The name a caller uses, and the name that is reported. */
  name: string;
  /** Width and height in twips, portrait. */
  width: number;
  height: number;
  /** The printer paper code that Word writes in `w:code`, so that its page setup dialog names the paper. Absent when no code is certain. */
  code?: number;
  /** The margins that a new document of this paper gets when the caller gives none. */
  margins: PageMargins;
  /** The size as it is written in the docs. */
  label: string;
}

const inch = (n: number): number => Math.round(n * 1440);
const mm = (n: number): number => Math.round((n / 25.4) * 1440);
const margins = (top: number, side: number, distance: number): PageMargins => ({ top, bottom: top, left: side, right: side, header: distance, footer: distance });

const LETTER_MARGINS = margins(inch(1), inch(1), inch(0.5));
const MID_MARGINS = margins(inch(0.75), inch(0.75), inch(0.4));
// The A4 values are what the A4 templates have always had: 1 in above and below, 25 mm at the sides.
const LARGE_METRIC_MARGINS = margins(1440, 1417, 709);
const A5_MARGINS = margins(mm(20), mm(20), mm(10));
const A6_MARGINS = margins(mm(15), mm(15), mm(7.5));

export const PAPERS: readonly Paper[] = [
  { name: 'Letter', width: inch(8.5), height: inch(11), code: 1, margins: LETTER_MARGINS, label: '8.5 x 11 in' },
  { name: 'Legal', width: inch(8.5), height: inch(14), code: 5, margins: LETTER_MARGINS, label: '8.5 x 14 in' },
  { name: 'Tabloid', width: inch(11), height: inch(17), code: 3, margins: LETTER_MARGINS, label: '11 x 17 in' },
  { name: 'Executive', width: inch(7.25), height: inch(10.5), code: 7, margins: MID_MARGINS, label: '7.25 x 10.5 in' },
  { name: 'Statement', width: inch(5.5), height: inch(8.5), code: 6, margins: MID_MARGINS, label: '5.5 x 8.5 in' },
  { name: 'Folio', width: inch(8.5), height: inch(13), code: 14, margins: LETTER_MARGINS, label: '8.5 x 13 in' },
  // Word writes 11906 x 16838 for A4 (210 x 297 mm), not the rounding of 11905.5.
  { name: 'A3', width: 16838, height: 23811, code: 8, margins: LARGE_METRIC_MARGINS, label: '297 x 420 mm' },
  { name: 'A4', width: 11906, height: 16838, code: 9, margins: LARGE_METRIC_MARGINS, label: '210 x 297 mm' },
  { name: 'A5', width: mm(148), height: mm(210), code: 11, margins: A5_MARGINS, label: '148 x 210 mm' },
  { name: 'A6', width: mm(105), height: mm(148), code: 70, margins: A6_MARGINS, label: '105 x 148 mm' },
  { name: 'B4', width: mm(250), height: mm(353), code: 42, margins: LARGE_METRIC_MARGINS, label: '250 x 353 mm' },
  { name: 'B5', width: mm(176), height: mm(250), margins: A5_MARGINS, label: '176 x 250 mm' },
  { name: 'B6', width: mm(125), height: mm(176), margins: A6_MARGINS, label: '125 x 176 mm' },
];

/** The names of the named papers, in table order. */
export const PAPER_NAMES: readonly string[] = PAPERS.map((p) => p.name);

/** A named paper, found without regard to case, or undefined. */
export function findPaper(name: string): Paper | undefined {
  const key = name.trim().toLowerCase();
  return PAPERS.find((p) => p.name.toLowerCase() === key);
}

/**
 * The named paper that a width and height match, in either orientation and within one twip, or undefined for a custom size.
 */
export function detectPaper(width: number, height: number): Paper | undefined {
  const near = (a: number, b: number): boolean => Math.abs(a - b) <= 1;
  return PAPERS.find((p) => (near(width, p.width) && near(height, p.height)) || (near(width, p.height) && near(height, p.width)));
}

/**
 * The margins that a new document gets for a page with this width and height and no margins given. A named paper has its own.
 * A custom size gets 1 in when its short side is 8 in or more, and scales down with the short side below that, never under 0.5 in.
 */
export function defaultMarginsFor(width: number, height: number): PageMargins {
  const paper = detectPaper(width, height);
  if (paper) return { ...paper.margins };
  const short = Math.min(width, height);
  const side = Math.max(inch(0.5), Math.min(inch(1), Math.round((short / inch(8)) * inch(1))));
  // The text area must keep its minimum on each axis, so a small page gets smaller margins rather than none left for text.
  const across = Math.max(0, Math.min(side, Math.floor((width - MIN_TEXT_WIDTH) / 2)));
  const down = Math.max(0, Math.min(side, Math.floor((height - MIN_TEXT_HEIGHT) / 2)));
  const distance = Math.round(side / 2);
  return { top: down, bottom: down, left: across, right: across, header: Math.min(distance, down), footer: Math.min(distance, down) };
}
