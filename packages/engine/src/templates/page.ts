import { checkPage, defaultMarginsFor, findPaper, resolveMargins, resolvePageSize, type MarginsInput, type PageSizeInput } from '../page';
import { toError } from '../errors';
import { LETTER, type PageSetup } from './lib';

/** The page of a new document, as a caller writes it. Everything is optional: no page means US Letter, portrait, 1 in margins. */
export interface NewPageOptions {
  /** A paper name ("A4", "Letter", ...) or a size ("148x210mm", { width, height }). Default Letter. */
  size?: PageSizeInput;
  /** Turns the page. A named paper is portrait unless this says landscape; a given size is turned to match. */
  orientation?: 'portrait' | 'landscape';
  /** One length for the four page margins, or any of top, bottom, left, right, header, footer, gutter. */
  margins?: MarginsInput;
}

/** The narrowest text width, in twips, that a template's tables can use. Below it the template is refused; every template other than blank warns below `WARN_BELOW`. */
const REFUSE_BELOW: Record<string, number> = { invoice: 5760, proposal: 5760 };
const WARN_BELOW = 4320;

/**
 * Resolves the page of a new document to twips and checks it. Returns the page and warnings. Throws `E_VALIDATION` for a size or margins that
 * cannot be laid out, and for the old `pageSize` option, which was replaced by `page`.
 */
export function resolveNewPage(template: string, options: unknown): { page: PageSetup; warnings: string[] } {
  const given = options === undefined || options === null ? {} : options;
  if (typeof given !== 'object' || Array.isArray(given)) throw toError('E_VALIDATION', 'the page option must be an object', 'pass { size, orientation, margins }');
  const o = given as Record<string, unknown>;
  for (const key of Object.keys(o)) {
    if (key === 'size' || key === 'orientation' || key === 'margins') continue;
    throw toError('E_VALIDATION', `unknown page field '${key}'`, 'the fields are size, orientation and margins');
  }
  if (o.orientation !== undefined && o.orientation !== 'portrait' && o.orientation !== 'landscape') {
    throw toError('E_VALIDATION', `orientation must be "portrait" or "landscape", got ${JSON.stringify(o.orientation)}`, 'use "portrait" or "landscape"');
  }
  const resolved = o.size === undefined ? { width: LETTER.width, height: LETTER.height, paper: findPaper('Letter'), named: true } : resolvePageSize(o.size, 'page.size');
  let { width, height } = resolved;
  if (o.orientation === 'landscape' && width < height) [width, height] = [height, width];
  if (o.orientation === 'portrait' && width > height) [width, height] = [height, width];
  const base = resolved.paper ? resolved.paper.margins : defaultMarginsFor(width, height);
  const given2 = o.margins === undefined ? {} : resolveMargins(o.margins, 'page.margins');
  const page: PageSetup = { width, height, top: base.top, bottom: base.bottom, left: base.left, right: base.right, header: base.header, footer: base.footer };
  for (const [k, v] of Object.entries(given2)) (page as unknown as Record<string, number>)[k] = v as number;
  if (resolved.paper?.code !== undefined) page.code = resolved.paper.code;
  // A named paper brings its own margins; when the caller gave only some, the header and footer distance must still fit the margin they sit in.
  const where = `page for the ${template} template`;
  const warnings = checkPage(page, where);
  const text = width - page.left - page.right - (page.gutter ?? 0);
  const refuse = REFUSE_BELOW[template];
  if (refuse !== undefined && text < refuse) {
    throw toError('E_VALIDATION', `${where}: the text is ${Math.round((text / 1440) * 100) / 100} in wide, which is too narrow for the ${template} template's tables (at least ${refuse / 1440} in)`, `use a wider page or smaller margins, or start from the blank template`);
  }
  if (refuse === undefined && template !== 'blank' && text < WARN_BELOW) {
    warnings.push(`${where}: the text is ${Math.round((text / 1440) * 100) / 100} in wide; the ${template} template's tables and headings may wrap awkwardly`);
  }
  return { page, warnings };
}
