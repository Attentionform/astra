/**
 * Drawing ids are numbered the same way on every pass. A drawing nested in a text box owns its own
 * ids, so the outer drawing does not renumber them again.
 */
import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { dedupeDrawingDocPrIds } from '@attentionform/astra-engine';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

/** A drawing whose docPr and cNvPr carry `id`; `inner` is placed inside it (as a text box would). */
function drawing(id: string | null, inner = ''): string {
  const attr = id === null ? '' : ` id="${id}"`;
  return `<w:drawing><wp:inline><wp:docPr${attr} name="d"/><pic:pic><pic:nvPicPr><pic:cNvPr${attr} name="p"/></pic:nvPicPr></pic:pic>${inner}</wp:inline></w:drawing>`;
}

function parse(body: string, root = 'w:body'): Document {
  return new DOMParser().parseFromString(`<${root} ${NS}><w:p><w:r>${body}</w:r></w:p></${root}>`, 'text/xml');
}

const xml = (doc: Document): string => new XMLSerializer().serializeToString(doc);
const ids = (doc: Document): string[] => [
  ...[...xml(doc).matchAll(/<(?:wp:docPr|pic:cNvPr)[^>]*\sid="(\d+)"/g)].map((m) => m[1]),
];

describe('dedupeDrawingDocPrIds', () => {
  it('leaves a nested drawing with valid ids alone, on every pass', () => {
    const doc = parse(drawing('1', drawing('2')));
    dedupeDrawingDocPrIds(doc);
    const once = xml(doc);
    expect(ids(doc)).toEqual(['1', '1', '2', '2']);
    dedupeDrawingDocPrIds(doc);
    expect(xml(doc)).toBe(once);
  });

  it('numbers duplicate ids uniquely in one pass and changes nothing in the next', () => {
    const doc = parse(drawing('1', drawing('1')) + drawing('1'));
    dedupeDrawingDocPrIds(doc);
    const once = xml(doc);
    const pairs = ids(doc);
    expect(new Set([pairs[0], pairs[2], pairs[4]]).size).toBe(3);
    dedupeDrawingDocPrIds(doc);
    expect(xml(doc)).toBe(once);
  });

  it('numbers a drawing without an id once', () => {
    const doc = parse(drawing('5') + drawing(null));
    dedupeDrawingDocPrIds(doc);
    const once = xml(doc);
    expect(ids(doc)).toEqual(['5', '5', '6', '6']);
    dedupeDrawingDocPrIds(doc);
    expect(xml(doc)).toBe(once);
  });

  it('follows the same rule in a header part', () => {
    const doc = parse(drawing('3', drawing('4')), 'w:hdr');
    dedupeDrawingDocPrIds(doc);
    const once = xml(doc);
    expect(ids(doc)).toEqual(['3', '3', '4', '4']);
    dedupeDrawingDocPrIds(doc);
    expect(xml(doc)).toBe(once);
  });
});
