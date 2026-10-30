import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { assertNoXxe } from './ooxml/xxe-guard';

export class EquationValidationError extends Error { constructor(m: string) { super(m); this.name = 'EquationValidationError'; } }
export class CitationValidationError extends Error { constructor(m: string) { super(m); this.name = 'CitationValidationError'; } }
export class FormattingValidationError extends Error { constructor(m: string) { super(m); this.name = 'FormattingValidationError'; } }
export class RunCountValidationError extends Error { constructor(m: string) { super(m); this.name = 'RunCountValidationError'; } }
export class NamespaceValidationError extends Error { constructor(m: string) { super(m); this.name = 'NamespaceValidationError'; } }
export class TrackedChangeValidationError extends Error { constructor(m: string) { super(m); this.name = 'TrackedChangeValidationError'; } }

const parser = new DOMParser({
  errorHandler: {
    warning: (w) => { throw new Error(w) },
    error: (e) => { throw new Error(e) },
    fatalError: (f) => { throw new Error(f) }
  }
});
const serializer = new XMLSerializer();

function extractNamespaceUris(xmlStr: string): Set<string> {
  const uris = new Set<string>();
  const re = /xmlns(?::[a-zA-Z0-9_]+)?="([^"]+)"/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xmlStr)) !== null) {
    uris.add(m[1]);
  }
  return uris;
}

// Standard OOXML URIs that are always allowed even if absent from the original
// node (alias changes and re-declarations of the same URI are harmless).
const STANDARD_OOXML_URIS = new Set<string>([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://schemas.openxmlformats.org/officeDocument/2006/math',
  'urn:schemas-microsoft-com:vml',
  'urn:schemas-microsoft-com:office:office',
  'urn:schemas-microsoft-com:office:word',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://schemas.openxmlformats.org/drawingml/2006/chart',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://schemas.microsoft.com/office/word/2010/wordml',
  'http://schemas.microsoft.com/office/word/2012/wordml',
  'http://schemas.openxmlformats.org/package/2006/relationships',
  'http://schemas.openxmlformats.org/package/2006/content-types',
  'http://www.w3.org/2000/xmlns/',
  'http://www.w3.org/XML/1998/namespace',
]);

export interface ValidationOptions {
  /** Skip Layer 3 (pPr preservation) — needed when applying style changes */
  allowStyleChanges?: boolean;
  /** Skip Layer 4 (run count check) — needed when inserting/removing text runs */
  allowRunCountChanges?: boolean;
  /** Maximum allowed run count difference (default: 10) */
  maxRunCountDelta?: number;
}

export function sanitizeXml(xmlStr: string): string {
  if (!xmlStr || typeof xmlStr !== 'string') return '';
  assertNoXxe(xmlStr, 'XML part');
  if (/<script/i.test(xmlStr)) {
    throw new Error("Security Error: Executable script tags are prohibited in OOXML.");
  }
  return xmlStr;
}

export function validateNode(originalXml: string, newXml: string, options: ValidationOptions = {}): boolean {
  sanitizeXml(originalXml);
  sanitizeXml(newXml);

  let origDom: Document;
  let newDom: Document;

  try {
    origDom = parser.parseFromString(originalXml, 'text/xml');
    newDom = parser.parseFromString(newXml, 'text/xml');
  } catch {
    throw new Error("Invalid XML formatting in LLM output");
  }

  const origNode = origDom.documentElement;
  void origNode;
  const newNode = newDom.documentElement;
  void newNode;

  // Layer 1: Equation counts
  const origMath = origDom.getElementsByTagName('m:oMath').length;
  const newMath = newDom.getElementsByTagName('m:oMath').length;
  if (origMath !== newMath) {
    throw new EquationValidationError(`Equation count mismatch: original has ${origMath}, new has ${newMath}`);
  }

  // Layer 2: Citation counts
  const origCite = origDom.getElementsByTagName('w:fldSimple').length;
  const newCite = newDom.getElementsByTagName('w:fldSimple').length;
  if (origCite !== newCite) {
    throw new CitationValidationError(`Citation count mismatch: original has ${origCite}, new has ${newCite}`);
  }

  // Layer 3: Formatting (pPr) preservation — SKIPPED when allowStyleChanges is true
  if (!options.allowStyleChanges) {
    const origPPr = origDom.getElementsByTagName('w:pPr')[0];
    const newPPr = newDom.getElementsByTagName('w:pPr')[0];
    const origPPrStr = origPPr ? serializer.serializeToString(origPPr) : '';
    const newPPrStr = newPPr ? serializer.serializeToString(newPPr) : '';
    if (origPPrStr !== newPPrStr) {
      throw new FormattingValidationError(`Formatting <w:pPr> was altered. LLM is not allowed to change styles.`);
    }
  }

  // Layer 4: Run count sanity check — SKIPPED when allowRunCountChanges is true
  if (!options.allowRunCountChanges) {
    const maxDelta = options.maxRunCountDelta ?? 10;
    const origRuns = origDom.getElementsByTagName('w:r').length;
    const newRuns = newDom.getElementsByTagName('w:r').length;
    if (Math.abs(origRuns - newRuns) > maxDelta) {
      throw new RunCountValidationError(`Run count changed drastically (from ${origRuns} to ${newRuns}). This risks breaking inline styling.`);
    }
  }

  // Layer 5: Namespace preservation — compare URI sets, ignoring prefix aliases.
  const origNs = extractNamespaceUris(originalXml);
  const newNs = extractNamespaceUris(newXml);
  for (const uri of newNs) {
    if (!origNs.has(uri) && !STANDARD_OOXML_URIS.has(uri)) {
      throw new NamespaceValidationError(`New unexpected namespace introduced: ${uri}`);
    }
  }

  // Layer 6: Tracked changes preservation — existing revisions must survive an
  // edit, but ADDING new ones (w:ins/w:del) is a legitimate authoring action
  // (manual XML editor, trackedChange workflows). Only loss is a violation.
  const origIns = origDom.getElementsByTagName('w:ins').length;
  const origDel = origDom.getElementsByTagName('w:del').length;
  const newIns = newDom.getElementsByTagName('w:ins').length;
  const newDel = newDom.getElementsByTagName('w:del').length;
  if (newIns < origIns || newDel < origDel) {
    throw new TrackedChangeValidationError(`Tracked changes must not be removed. Original ins/del: ${origIns}/${origDel}, New ins/del: ${newIns}/${newDel}. Accept or reject them instead.`);
  }

  return true;
}

