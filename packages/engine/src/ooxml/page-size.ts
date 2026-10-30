/**
 * Page dimensions for `w:pgSz`. Word and LibreOffice lay a page out from
 * `w:w` and `w:h`; the `w:orient` flag alone does not turn it, so the two must
 * always agree.
 */

export interface PageDimensions {
  width: number;
  height: number;
}

/** What a section without a `w:pgSz` size is laid out as (the blank template's page). */
export const DEFAULT_PAGE_DIMENSIONS: PageDimensions = { width: 12240, height: 15840 };

export type PageOrientation = 'portrait' | 'landscape';

/** Width and height as the section holds them; a missing or non-positive value falls back to Letter. */
export function readPageDimensions(pgSz: Element): PageDimensions {
  const width = parseInt(pgSz.getAttribute('w:w') ?? '', 10);
  const height = parseInt(pgSz.getAttribute('w:h') ?? '', 10);
  return {
    width: width > 0 ? width : DEFAULT_PAGE_DIMENSIONS.width,
    height: height > 0 ? height : DEFAULT_PAGE_DIMENSIONS.height,
  };
}

/** The orientation a section is laid out in: the flag wins, otherwise the longer side does. */
export function readPageOrientation(pgSz: Element): PageOrientation {
  const flag = pgSz.getAttribute('w:orient');
  if (flag === 'landscape' || flag === 'portrait') return flag;
  const { width, height } = readPageDimensions(pgSz);
  return width > height ? 'landscape' : 'portrait';
}

/**
 * The same page turned to `orientation`: the longer side is the width for
 * landscape and the height for portrait. A square page is left alone.
 */
export function orientPageDimensions(dimensions: PageDimensions, orientation: PageOrientation): PageDimensions {
  const long = Math.max(dimensions.width, dimensions.height);
  const short = Math.min(dimensions.width, dimensions.height);
  return orientation === 'landscape' ? { width: long, height: short } : { width: short, height: long };
}
