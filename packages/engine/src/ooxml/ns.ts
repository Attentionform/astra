import { resolveReproducibleMode } from '../zip-hardening';

export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const WORDPROCESSING_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
export const WP_NS = WORDPROCESSING_DRAWING_NS;
export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

export function createWordprocessingDrawingElement(doc: Document, tagName: string): Element {
  const qualifiedName = tagName.includes(':') ? tagName : `wp:${tagName}`;
  return doc.createElementNS(WORDPROCESSING_DRAWING_NS, qualifiedName);
}

export function createWordElement(doc: Document, tagName: string): Element {
  const qualifiedName = tagName.includes(':') ? tagName : `w:${tagName}`;
  return doc.createElementNS(WORD_NS, qualifiedName);
}

export function createDrawingElement(doc: Document, tagName: string): Element {
  const qualifiedName = tagName.includes(':') ? tagName : `a:${tagName}`;
  return doc.createElementNS(DRAWING_NS, qualifiedName);
}

export function createChartElement(doc: Document, tagName: string): Element {
  const qualifiedName = tagName.includes(':') ? tagName : `c:${tagName}`;
  return doc.createElementNS(CHART_NS, qualifiedName);
}

export function createRelationshipsElement(doc: Document, tagName: string): Element {
  return doc.createElementNS(RELATIONSHIPS_NS, tagName);
}

export function createContentTypesElement(doc: Document, tagName: string): Element {
  return doc.createElementNS(CONTENT_TYPES_NS, tagName);
}

export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

export function createMathElement(doc: Document, tagName: string): Element {
  const qualifiedName = tagName.includes(':') ? tagName : `m:${tagName}`;
  return doc.createElementNS(MATH_NS, qualifiedName);
}

// A-2: process-wide high-water mark for drawing ids (wp:docPr / pic:cNvPr).
// Module counters that reset per process collide with ids already in the
// document (fresh-process insert into T5 re-mints id 1 -> Sem_UniqueAttributeValue).
let drawingDocPrHighWater = 0;

/**
 * A-2: allocates the next free drawing id by scanning the live document DOM
 * for all wp:docPr/pic:cNvPr ids and returning max+1 (floored by the
 * process-wide high-water mark so ids minted but not yet in the DOM can
 * never be re-issued). Shared by image insert and chart insert, which occupy
 * the same id space. Deterministic default: DOM scan only (no high-water),
 * so same input twice mints the same id; the high-water reservation applies
 * only when opted out of deterministic output.
 */
export function allocateNextDrawingDocPrId(doc?: Document | null): number {
  let repro = true;
  try {
    repro = resolveReproducibleMode(undefined);
  } catch {
    repro = true;
  }
  let maxId = repro ? 0 : drawingDocPrHighWater;
  if (doc) {
    for (const tag of ['wp:docPr', 'pic:cNvPr']) {
      const els = doc.getElementsByTagName(tag);
      for (let i = 0; i < els.length; i++) {
        const raw = parseInt((els[i] as unknown as Element).getAttribute('id') || '0', 10);
        if (!isNaN(raw) && raw > maxId) maxId = raw;
      }
    }
  }
  drawingDocPrHighWater = maxId + 1;
  if (drawingDocPrHighWater < 1) drawingDocPrHighWater = 1;
  return drawingDocPrHighWater;
}
