import { XMLSerializer } from '@xmldom/xmldom';
import { DocumentAST } from '../ast';
import { BaseOperation } from './base';
import { DocxParser, upsertSectPrChild } from '../parser';
import { sectionPosition } from './section-number';
import { orientPageDimensions, readPageDimensions, readPageOrientation } from '../ooxml/page-size';
import { pageSetupFromSectPr } from '../parser/sections';
import { DOMParser } from '../ooxml/dom-parser';
import { assertNoXxe } from '../ooxml/xxe-guard';
import { SETTINGS_CHILD_ORDER, sortChildrenBySchema } from '../ooxml/schema-order';
import { createWordElement } from '../ooxml/ns';
import { toError } from '../errors';
import { checkPage, defaultMarginsFor, detectPaper, resolveMargins, resolvePageSize, type MarginName, type MarginsInput, type PageGeometry, type PageSizeInput } from '../page';

/**
 * The fields of a page setup. Every field is optional, and a field that is left out keeps what the section has.
 * Lengths are twips (1440 to an inch) when they are numbers, and may be strings with a unit, such as `"25mm"` or `"1in"`.
 */
export type PageSetupFields = {
  /** A paper name (`"A4"`, `"Letter"`, `"A5"` and the others in `PAPER_NAMES`), a width and a height (`{ width: "148mm", height: "210mm" }`), or a size string (`"6x9in"`). */
  paperSize?: PageSizeInput;
  orientation?: 'portrait' | 'landscape';
  /** One length for the four page margins, or an object with any of `top`, `bottom`, `left`, `right`, `header`, `footer` and `gutter`. */
  margins?: MarginsInput;
  /** The edge that the gutter is on. The setting is for the whole document, as Word keeps it. */
  gutterPosition?: 'left' | 'top';
  /** Mirror the inner and outer margins for facing pages. The setting is for the whole document, as Word keeps it. */
  mirrorMargins?: boolean;
  columns?: number;
  /** The section to change, counted from 1. Without it every section changes. */
  sectionIndex?: number;
};

/**
 * Set the paper size, the orientation, the margins and the columns of a section (`action: 'pageSetup'`). The fields can be given at the top level or in `params`.
 */
export interface PageSetupOperation extends BaseOperation, PageSetupFields {
  action: 'pageSetup';
  params?: PageSetupFields;
}

/** What a page setup reports. `warnings` are advice, such as a page that Word cannot read. */
export interface PageSetupResult {
  warnings: string[];
}

const SETTINGS_PART = 'word/settings.xml';

const intAttr = (el: Element | undefined, name: string): number | undefined => {
  const v = parseInt(el?.getAttribute(name) ?? '', 10);
  return Number.isFinite(v) ? v : undefined;
};

/** True when the settings part switches `tag` on (the element is present and not `w:val="0"`, `false` or `off`). */
function settingsFlag(settingsXml: string | null, tag: string): boolean {
  if (settingsXml === null) return false;
  const m = new RegExp(`<${tag}(\\s[^>]*)?/?>`).exec(settingsXml);
  if (!m) return false;
  return !/w:val="(0|false|off)"/.test(m[1] ?? '');
}

/** Switches the document-wide margin settings. Needs the settings part, which every Word document has. */
function writeSettings(parser: DocxParser, changes: { mirrorMargins?: boolean; gutterAtTop?: boolean }): void {
  const zip = parser.getZipAdapter();
  const raw = zip.readTextOrNull(SETTINGS_PART);
  if (raw === null || raw.trim() === '') {
    throw toError('E_VALIDATION', 'pageSetup cannot set mirrorMargins or gutterPosition: the document has no word/settings.xml', 'open the file in Word and save it again, or compose into a document made by ASTra');
  }
  assertNoXxe(raw, SETTINGS_PART);
  const doc = new DOMParser().parseFromString(raw, 'text/xml') as unknown as Document;
  const root = doc.documentElement as unknown as Element;
  if (!root || root.tagName !== 'w:settings') {
    throw toError('E_VALIDATION', 'pageSetup cannot read word/settings.xml: its root is not w:settings', 'restore word/settings.xml from a clean copy of the file');
  }
  const set = (tag: string, on: boolean): void => {
    const existing = Array.from(root.childNodes).filter((c) => c.nodeType === 1 && (c as Element).tagName === tag) as Element[];
    for (const e of existing) root.removeChild(e);
    if (on) root.appendChild(createWordElement(doc, tag));
  };
  if (changes.mirrorMargins !== undefined) set('w:mirrorMargins', changes.mirrorMargins);
  if (changes.gutterAtTop !== undefined) set('w:gutterAtTop', changes.gutterAtTop);
  sortChildrenBySchema(root, SETTINGS_CHILD_ORDER);
  zip.writeText(SETTINGS_PART, new XMLSerializer().serializeToString(doc as never));
}

interface SectionPlan {
  sectPr: Element;
  width: number;
  height: number;
  margins: { top: number; bottom: number; left: number; right: number; header: number; footer: number; gutter: number };
  orient: 'portrait' | 'landscape' | null;
  /** The printer paper code for `w:code`: a number to write, null to remove it, undefined to leave it. */
  code: number | null | undefined;
}

/** The names of the fields a page setup takes. */
const FIELD_NAMES = ['paperSize', 'orientation', 'margins', 'gutterPosition', 'mirrorMargins', 'columns', 'sectionIndex'] as const;

/** Names that an earlier version took, and what replaces each. They are refused, never ignored. */
const REPLACED: Record<string, string> = {
  marginTop: 'margins.top',
  marginBottom: 'margins.bottom',
  marginLeft: 'margins.left',
  marginRight: 'margins.right',
  headerDistance: 'margins.header',
  footerDistance: 'margins.footer',
  header: 'margins.header',
  footer: 'margins.footer',
  margin: 'margins',
  pageSize: 'paperSize (a paper name, a size such as "148x210mm", or { width, height })',
};

/**
 * Refuses a field that an earlier version took, and a field that is not a page setup field, before anything is read. An unknown name must
 * never be dropped: a request that was ignored would report success with the page unchanged.
 */
function assertKnownFields(op: PageSetupOperation): void {
  const refuse = (where: string, key: string, known: boolean): never => {
    const replacement = REPLACED[key];
    if (replacement) {
      throw toError('E_VALIDATION', `pageSetup${where}: '${key}' is no longer a field`, `use ${replacement} instead, for example { "margins": { "top": "20mm", "left": "1in" } }`);
    }
    throw toError('E_VALIDATION', `pageSetup${where}: '${key}' is not a page setup field`, `use ${FIELD_NAMES.join(', ')}${known ? '' : ''}`);
  };
  for (const key of Object.keys(op)) {
    if (key in REPLACED && !FIELD_NAMES.includes(key as never)) refuse('', key, true);
  }
  if (op.params !== undefined && op.params !== null) {
    if (typeof op.params !== 'object' || Array.isArray(op.params)) {
      throw toError('E_VALIDATION', 'pageSetup: params must be an object', 'pass the page setup fields in params, or at the top level');
    }
    for (const key of Object.keys(op.params)) {
      if (!(FIELD_NAMES as readonly string[]).includes(key)) refuse(' params', key, false);
    }
  }
}

/** The settings that make `w:gutterAtTop` invalid when present (ISO/IEC 29500, 17.15.1.50). */
const GUTTER_AT_TOP_CONFLICTS: Array<[string, string]> = [
  ['w:mirrorMargins', 'mirrorMargins'],
  ['w:bookFoldPrinting', 'book fold printing'],
  ['w:bookFoldRevPrinting', 'reverse book fold printing'],
  ['w:printTwoOnOne', 'two pages per sheet'],
];

/**
 * Changes the section's <w:sectPr> (w:pgSz / w:pgMar / w:cols) for real, with schema-correct child order, and keeps the AST
 * section metadata in sync so readers see the change. Every section that the request affects is checked before any is changed,
 * so a refusal changes nothing. A setting that is for the whole document (the gutter edge) is checked against every section.
 */
export function executePageSetup(op: PageSetupOperation, ast: DocumentAST, parser?: DocxParser): PageSetupResult {
  const sectPrs = parser?.getSectPrElements() ?? [];
  if (sectPrs.length === 0 || !parser) {
    throw new Error('pageSetup requires a document section (<w:sectPr>); attach a DocxParser to the engine.');
  }
  assertKnownFields(op);

  const f: PageSetupFields = { ...op, ...(op.params ?? {}) };
  const size = f.paperSize !== undefined ? resolvePageSize(f.paperSize) : undefined;
  const given = resolveMargins(f.margins);
  const givenNames = Object.keys(given) as MarginName[];
  const columns = f.columns;
  const sectionIndex = sectionPosition('pageSetup', f.sectionIndex, sectPrs.length, 'omit it to change every section');
  if (f.orientation !== undefined && f.orientation !== 'portrait' && f.orientation !== 'landscape') {
    throw toError('E_VALIDATION', `orientation '${String(f.orientation)}' is not portrait or landscape`, 'use portrait or landscape');
  }
  if (f.gutterPosition !== undefined && f.gutterPosition !== 'left' && f.gutterPosition !== 'top') {
    throw toError('E_VALIDATION', `gutterPosition '${String(f.gutterPosition)}' is not left or top`, 'use left or top');
  }
  if (f.mirrorMargins !== undefined && typeof f.mirrorMargins !== 'boolean') {
    throw toError('E_VALIDATION', `mirrorMargins must be true or false, got ${String(f.mirrorMargins)}`, 'use true or false');
  }
  if (columns !== undefined && (!Number.isInteger(columns) || columns < 1)) {
    throw toError('E_VALIDATION', `columns must be a whole number of 1 or more, got ${String(columns)}`, 'give the number of text columns, for example 2');
  }

  const settingsXml = parser.getZipAdapter().readTextOrNull(SETTINGS_PART);
  const gutterAtTopBefore = settingsFlag(settingsXml, 'w:gutterAtTop');
  const gutterAtTop = f.gutterPosition !== undefined ? f.gutterPosition === 'top' : gutterAtTopBefore;
  const mirror = f.mirrorMargins !== undefined ? f.mirrorMargins : settingsFlag(settingsXml, 'w:mirrorMargins');

  // A top gutter must not be used with mirrored margins, book fold or two pages per sheet. Only a request that sets one of the two refuses;
  // a file that already combines them can still be edited for something else.
  if ((f.gutterPosition !== undefined || f.mirrorMargins !== undefined) && gutterAtTop) {
    const clash = GUTTER_AT_TOP_CONFLICTS.filter(([tag]) => (tag === 'w:mirrorMargins' ? mirror : settingsFlag(settingsXml, tag))).map(([, label]) => label);
    if (clash.length > 0) {
      throw toError('E_VALIDATION', `pageSetup: a top gutter cannot be used with ${clash.join(' or ')}`, 'put the gutter on the left (gutterPosition: "left"), or switch the other setting off in the same request');
    }
  }

  const changesPage = size !== undefined || f.orientation !== undefined || givenNames.length > 0;
  // The gutter edge is a setting of the whole document, so it moves the text area of every section that has a gutter.
  const changesAllSections = f.gutterPosition !== undefined && gutterAtTop !== gutterAtTopBefore;

  const warnings: string[] = [];
  const plans: SectionPlan[] = [];
  sectPrs.forEach((sectPr, i) => {
    const isTarget = sectionIndex === undefined || sectionIndex === i;
    if (!(isTarget && changesPage) && !changesAllSections) return;
    const where = sectPrs.length === 1 ? 'page setup' : `section ${i + 1}`;
    const pgSz = sectPr.getElementsByTagName('w:pgSz')[0];
    const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
    const keptOrientation = pgSz ? readPageOrientation(pgSz) : 'portrait';
    const before = pgSz ? readPageDimensions(pgSz) : { width: 12240, height: 15840 };
    let page = before;
    if (isTarget && size !== undefined) {
      // A paper name is a portrait page that turns to the orientation in use; a width and a height are the page as given.
      page = size.named ? orientPageDimensions({ width: size.width, height: size.height }, f.orientation ?? keptOrientation) : { width: size.width, height: size.height };
    }
    if (isTarget && f.orientation !== undefined) page = orientPageDimensions(page, f.orientation);
    const defaults = defaultMarginsFor(page.width, page.height);
    const current = {
      top: intAttr(pgMar, 'w:top') ?? defaults.top,
      bottom: intAttr(pgMar, 'w:bottom') ?? defaults.bottom,
      left: intAttr(pgMar, 'w:left') ?? defaults.left,
      right: intAttr(pgMar, 'w:right') ?? defaults.right,
      header: intAttr(pgMar, 'w:header') ?? defaults.header,
      footer: intAttr(pgMar, 'w:footer') ?? defaults.footer,
      gutter: intAttr(pgMar, 'w:gutter') ?? 0,
    };
    const margins = isTarget ? { ...current, ...given } : current;
    const resulting: PageGeometry = { ...page, ...margins, gutterAtTop };
    const previous: PageGeometry = { ...before, ...current, gutterAtTop: gutterAtTopBefore };
    for (const w of checkPage(resulting, where, { size: isTarget && size !== undefined, margins: isTarget ? givenNames : [], before: previous })) if (!warnings.includes(w)) warnings.push(w);
    if (!(isTarget && changesPage)) return;
    const turned = page.width > page.height ? 'landscape' : 'portrait';
    const hadOrient = pgSz?.getAttribute('w:orient');
    plans.push({
      sectPr,
      width: page.width,
      height: page.height,
      margins,
      orient: f.orientation !== undefined ? f.orientation : size !== undefined && (turned === 'landscape' || hadOrient) ? turned : null,
      code: size !== undefined ? detectPaper(page.width, page.height)?.code ?? null : undefined,
    });
  });

  const targets = sectionIndex === undefined ? sectPrs : [sectPrs[sectionIndex]];
  for (const plan of plans) {
    const pgMar = upsertSectPrChild(plan.sectPr, 'w:pgMar');
    pgMar.setAttribute('w:top', String(plan.margins.top));
    pgMar.setAttribute('w:right', String(plan.margins.right));
    pgMar.setAttribute('w:bottom', String(plan.margins.bottom));
    pgMar.setAttribute('w:left', String(plan.margins.left));
    pgMar.setAttribute('w:header', String(plan.margins.header));
    pgMar.setAttribute('w:footer', String(plan.margins.footer));
    pgMar.setAttribute('w:gutter', String(plan.margins.gutter));
    const pgSz = upsertSectPrChild(plan.sectPr, 'w:pgSz');
    pgSz.setAttribute('w:w', String(plan.width));
    pgSz.setAttribute('w:h', String(plan.height));
    if (plan.orient !== null) pgSz.setAttribute('w:orient', plan.orient);
    if (plan.code === null) pgSz.removeAttribute('w:code');
    else if (plan.code !== undefined) pgSz.setAttribute('w:code', String(plan.code));
  }

  if (columns !== undefined) {
    for (const sectPr of targets) upsertSectPrChild(sectPr, 'w:cols').setAttribute('w:num', String(columns));
  }

  if (f.mirrorMargins !== undefined || f.gutterPosition !== undefined) {
    writeSettings(parser, {
      ...(f.mirrorMargins !== undefined ? { mirrorMargins: f.mirrorMargins === true } : {}),
      ...(f.gutterPosition !== undefined ? { gutterAtTop: f.gutterPosition === 'top' } : {}),
    });
  }

  // Keep the AST section metadata in sync with the changed XML.
  const syncSections = sectionIndex === undefined || changesAllSections || f.gutterPosition !== undefined ? ast.sections : ast.sections.slice(sectionIndex, sectionIndex + 1);
  for (const section of syncSections) {
    const sectPr = sectPrs[ast.sections.indexOf(section)];
    const fresh = sectPr ? pageSetupFromSectPr(sectPr) : {};
    const merged = { ...(section.pageSetup ?? {}), ...fresh };
    if (fresh.columns === undefined) delete merged.columns;
    if (gutterAtTop) merged.gutterAtTop = true;
    else delete merged.gutterAtTop;
    section.pageSetup = merged;
    section.version += 1;
  }
  return { warnings };
}
