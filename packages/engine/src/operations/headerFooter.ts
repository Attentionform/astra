import { DocumentAST } from '../ast';
import { DocxParser } from '../parser';
import { toError } from '../errors';
import { sectionPosition } from './section-number';
import { BaseOperation } from './base';
import { HeaderFooterManipulator } from '../manipulators/header-footer';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

export interface HeaderFooterOperation extends BaseOperation {
  action: 'headerFooter';
  operation: 'set' | 'remove';
  targetType: 'header' | 'footer';
  pageType?: 'default' | 'first' | 'even';
  text?: string; // Content of the header/footer
  xmlContent?: string;
  sectionIndex?: number;
  alignment?: 'left' | 'center' | 'right';
  numberFormat?: string;
}

export function executeHeaderFooter(op: HeaderFooterOperation, ast: DocumentAST, parser?: DocxParser): void {
  const manipulator = new HeaderFooterManipulator(ast, parser);
  const pageType = op.pageType || 'default';
  const sectionCount = Math.max(parser?.getSectPrElements().length ?? 0, ast.sections.length, 1);
  const sectionIndex = sectionPosition('headerFooter', op.sectionIndex, sectionCount) ?? 0;

  if (op.operation === 'remove') {
    manipulator.removeHeaderFooter({
      targetType: op.targetType,
      pageType,
      sectionIndex
    });
  } else {
    manipulator.setHeaderFooter({
      targetType: op.targetType,
      pageType,
      text: op.text,
      xmlContent: op.xmlContent,
      sectionIndex,
      alignment: op.alignment,
      numberFormat: op.numberFormat
    });
  }
}

// Shared header/footer targeting for field + insert ops.
//
// A page-number field placed in the body renders once where it sits; placed
// in a footer part it repeats on every page by construction. These helpers
// let field/insert ops route content into header/footer parts instead of the
// body. Body remains the default; only an explicit target changes routing.

export type HeaderFooterTarget = 'body' | 'header' | 'footer';
export type HeaderFooterTargetPageType = 'default' | 'first' | 'even';

export interface HeaderFooterTargetResolution {
  target: HeaderFooterTarget;
  pageType: HeaderFooterTargetPageType;
  sectionIndex: number;
}

const TARGET_VALUES: HeaderFooterTarget[] = ['body', 'header', 'footer'];
const PAGE_TYPE_VALUES: HeaderFooterTargetPageType[] = ['default', 'first', 'even'];

function targetError(message: string, remedy: string): Error {
  // Single envelope: text stays byte-identical (`E_VALIDATION: ... Remedy:
  // ...`), while `.code`/`.remedy` ride along for machine mapping.
  return toError('E_VALIDATION', message, remedy);
}

/**
 * Strips session-only `data-astra-id` tracking attributes from an element
 * subtree (in place, never throws). Header/footer parts persist verbatim,
 * so any surviving id fails schema validation at save.
 */
function stripSessionIds(root: Element): void {
  try {
    const stack: Element[] = [root];
    while (stack.length > 0) {
      const el = stack.pop() as Element;
      try {
        if (typeof el.removeAttribute === 'function') el.removeAttribute('data-astra-id');
      } catch {
        /* best-effort per node */
      }
      const kids = el.childNodes;
      for (let i = 0; i < kids.length; i++) {
        if (kids[i].nodeType === 1) stack.push(kids[i] as Element);
      }
    }
  } catch {
    /* stripping never blocks the append */
  }
}

/**
 * Validates a header/footer target triple shared by field + insert ops.
 * Reads the already-extracted raw values (callers accept them flat or under
 * params). Undefined target means body. Anything else fails closed before
 * any mutation so a typo never lands content in the wrong store.
 */
export function resolveHeaderFooterTarget(
  opName: string,
  rawTarget: unknown,
  rawPageType?: unknown,
  rawSectionIndex?: unknown
): HeaderFooterTargetResolution {
  const target: HeaderFooterTarget = rawTarget === undefined || rawTarget === null || rawTarget === ''
    ? 'body'
    : String(rawTarget) as HeaderFooterTarget;
  if (!TARGET_VALUES.includes(target)) {
    throw targetError(
      `${opName} unknown target '${String(rawTarget)}'. Expected one of ${TARGET_VALUES.map((v) => `'${v}'`).join('|')}.`,
      `pass target:'body' (default), target:'header', or target:'footer' — e.g. { action:'field', operation:'insertPageNumber', target:'footer' }.`
    );
  }
  const pageType: HeaderFooterTargetPageType = rawPageType === undefined || rawPageType === null || rawPageType === ''
    ? 'default'
    : String(rawPageType) as HeaderFooterTargetPageType;
  if (!PAGE_TYPE_VALUES.includes(pageType)) {
    throw targetError(
      `${opName} unknown pageType '${String(rawPageType)}'. Expected one of 'default'|'first'|'even'.`,
      `pass pageType:'default' (every page), 'first' (first page only, needs titlePg), or 'even' (even pages).`
    );
  }
  let sectionIndex = 0;
  if (rawSectionIndex !== undefined && rawSectionIndex !== null && rawSectionIndex !== '') {
    const n = typeof rawSectionIndex === 'number' ? rawSectionIndex : Number(String(rawSectionIndex));
    if (!Number.isInteger(n) || n < 1) {
      throw targetError(
        `${opName} invalid sectionIndex '${String(rawSectionIndex)}'. Expected a positive integer (sections count from 1).`,
        `pass sectionIndex:1 for the first section; re-read the structure with inspect_document for more sections.`
      );
    }
    sectionIndex = n - 1; // callers count sections from 1; everything below counts from 0
  }
  return { target, pageType, sectionIndex };
}

export interface AppendToHeaderFooterOptions {
  targetType: 'header' | 'footer';
  pageType?: HeaderFooterTargetPageType;
  sectionIndex?: number;
}

/**
 * Appends already-built XML elements (w:p / w:tbl roots) to a header/footer
 * part, creating the part (rels, content-type, sectPr reference) when absent
 * and preserving existing part content when present. Returns the part
 * reference. Fails closed when there is no parser (parts live in the zip, so
 * a headless AST cannot persist them) or when the section does not exist.
 */
export function appendElementsToHeaderFooter(
  ast: DocumentAST,
  parser: DocxParser | undefined,
  elements: Element[],
  options: AppendToHeaderFooterOptions
): { rId: string; partPath: string } {
  const opName = 'header/footer append';
  const targetType = options.targetType;
  if (targetType !== 'header' && targetType !== 'footer') {
    throw targetError(
      `${opName} unknown targetType '${String(targetType)}'. Expected 'header' or 'footer'.`,
      `pass target:'header' or target:'footer'.`
    );
  }
  const pageType = options.pageType || 'default';
  if (pageType !== 'default' && pageType !== 'first' && pageType !== 'even') {
    throw targetError(
      `${opName} unknown pageType '${String(pageType)}'. Expected 'default', 'first' or 'even'.`,
      `pass pageType:'default', 'first', or 'even'.`
    );
  }
  const sectionIndex = options.sectionIndex || 0;
  if (!parser) {
    throw targetError(
      `${opName} to ${targetType} needs a parser-attached session (parts live in the document package).`,
      `run against an opened document (open_document first) instead of a headless AST, or use target:'body'.`
    );
  }
  if (!ast.sections[sectionIndex]) {
    throw targetError(
      `${opName} sectionIndex ${sectionIndex + 1} is out of range (document has ${ast.sections.length} section(s)).`,
      `pass sectionIndex in [1, ${ast.sections.length}], counting from 1; re-read the structure with inspect_document.`
    );
  }
  if (elements.length === 0) {
    throw targetError(
      `${opName} received no content elements.`,
      `pass at least one paragraph element to append.`
    );
  }

  // Header/footer parts persist verbatim: unlike word/document.xml (whose
  // session ids the save path strips), any `data-astra-id` here reaches the
  // schema validator and fails the save. Strip session-only tracking ids
  // from the appended subtrees up front so insert/field writes into these
  // parts save VALID instead of dying at the validator-gated save.
  for (const el of elements) {
    stripSessionIds(el);
  }

  const relType = targetType === 'header'
    ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
    : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
  const partRe = targetType === 'header' ? /(?:^|\/)header\d+\.xml$/ : /(?:^|\/)footer\d+\.xml$/;

  // Scope the part lookup to this section's sectPr so a first/even variant
  // never hijacks the default part (mirrors the manipulator's scoped lookup).
  let existingRId: string | null = null;
  try {
    const sectPrs = parser.getSectPrElements() ?? [];
    const sectPr = sectPrs.length > sectionIndex ? sectPrs[sectionIndex] : null;
    if (sectPr) {
      const refTag = targetType === 'header' ? 'w:headerReference' : 'w:footerReference';
      for (const ref of Array.from(sectPr.getElementsByTagName(refTag))) {
        if ((ref.getAttribute('w:type') || 'default') === pageType) {
          existingRId = ref.getAttribute('r:id') || (ref as Element).getAttribute('id');
          break;
        }
      }
    }
  } catch { existingRId = null; }

  const zip = parser.getZipAdapter();
  if (existingRId && ast.relationships[existingRId]) {
    const rel = ast.relationships[existingRId];
    if (rel.type === relType && partRe.test(rel.target)) {
      const partPath = rel.target.startsWith('word/') ? rel.target : `word/${rel.target.replace(/^\//, '')}`;
      if (zip.hasFile(partPath)) {
        const partDom = new DOMParser().parseFromString(zip.readText(partPath), 'text/xml');
        const root = partDom.documentElement;
        const expectedRoot = targetType === 'header' ? 'w:hdr' : 'w:ftr';
        if (!root || root.tagName !== expectedRoot) {
          throw targetError(
            `${opName} existing ${targetType} part '${partPath}' has unexpected root '<${root ? root.tagName : 'none'}>'.`,
            `repair or remove the corrupt ${targetType} part (or set it fresh via headerFooter set) and retry.`
          );
        }
        for (const el of elements) {
          let imported: Element = el;
          try {
            if (typeof (partDom as Document).importNode === 'function') {
              imported = (partDom as Document).importNode(el, true) as Element;
            }
          } catch {
            imported = el;
          }
          root.appendChild(imported);
        }
        zip.writeText(partPath, new XMLSerializer().serializeToString(partDom));
        const targetMap = targetType === 'header' ? ast.headers : ast.footers;
        targetMap[existingRId] = { type: pageType, xml: partDom.documentElement };
        return { rId: existingRId, partPath };
      }
    }
  }

  // No reusable part: build a full part document embedding the new nodes and
  // let the manipulator wire rels, content-type, sectPr reference, and map.
  const expectedRoot = targetType === 'header' ? 'w:hdr' : 'w:ftr';
  const shell =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<${expectedRoot} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `</${expectedRoot}>`;
  const shellDom = new DOMParser().parseFromString(shell, 'text/xml');
  const shellRoot = shellDom.documentElement;
  for (const el of elements) {
    let imported: Element = el;
    try {
      if (typeof (shellDom as Document).importNode === 'function') {
        imported = (shellDom as Document).importNode(el, true) as Element;
      }
    } catch {
      imported = el;
    }
    shellRoot.appendChild(imported);
  }
  const manipulator = new HeaderFooterManipulator(ast, parser);
  return manipulator.setHeaderFooter({
    targetType,
    pageType,
    xmlContent: new XMLSerializer().serializeToString(shellDom),
    sectionIndex
  });
}

