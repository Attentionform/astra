import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { REVISION_TAGS, MOVE_MARKER_TAGS, maxRevisionId, countRevisionElements, walkWordElements } from '../../packages/engine/src/ooxml/revision-ids';
import { countRevisions } from '../../packages/engine/src/ooxml/revisions';

/**
 * Counting the revision elements and finding the highest id walked the document once per element name (29 walks, 17 ms
 * on 3,000 paragraphs). They now walk it once. The old code is kept here as the oracle: the two must agree on every
 * document in the repository and on the odd shapes below.
 */
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function oldMaxRevisionId(document: Document): number {
  let max = 0;
  for (const tag of [...REVISION_TAGS, ...MOVE_MARKER_TAGS]) {
    for (const el of Array.from(document.getElementsByTagNameNS(W, tag))) {
      const value = el.getAttributeNS(W, 'id');
      if (value !== null && /^\d+$/.test(value)) {
        const n = Number(value);
        if (!Number.isSafeInteger(n)) return Infinity;
        max = Math.max(max, n);
      }
    }
  }
  return max;
}

function oldCount(documents: readonly Document[]): number {
  let count = 0;
  for (const doc of documents) for (const tag of REVISION_TAGS) count += doc.getElementsByTagNameNS(W, tag).length;
  return count;
}

const parse = (body: string): Document =>
  new DOMParser().parseFromString(`<w:document xmlns:w="${W}" xmlns:x="urn:other"><w:body>${body}</w:body></w:document>`, 'text/xml') as unknown as Document;

describe('one-pass scan of revision elements', () => {
  const A = 'w:author="a" w:date="2026-01-01T00:00:00Z"';
  const shapes: [string, string][] = [
    ['no revisions', '<w:p><w:r><w:t>x</w:t></w:r></w:p>'],
    ['ins and del', `<w:p><w:ins w:id="3" ${A}><w:r><w:t>a</w:t></w:r></w:ins><w:del w:id="9" ${A}><w:r><w:delText>b</w:delText></w:r></w:del></w:p>`],
    ['nested wrappers', `<w:p><w:hyperlink><w:ins w:id="41" ${A}><w:r><w:t>a</w:t></w:r></w:ins></w:hyperlink></w:p>`],
    ['format changes', `<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="7" ${A}><w:rPr/></w:rPrChange></w:rPr><w:t>a</w:t></w:r><w:pPr><w:pPrChange w:id="8" ${A}><w:pPr/></w:pPrChange></w:pPr></w:p>`],
    ['move markers count for the highest id, not for the count', `<w:p><w:moveFromRangeStart w:id="50" w:name="m"/><w:moveFrom w:id="51" ${A}><w:r><w:delText>a</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="50"/></w:p>`],
    ['table revisions', `<w:tbl><w:tr><w:trPr><w:del w:id="12" ${A}/></w:trPr><w:tc><w:tcPr><w:cellDel w:id="13" ${A}/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`],
    ['non-numeric and missing ids are ignored', `<w:p><w:ins w:id="abc" ${A}/><w:ins ${A}/><w:ins w:id="-4" ${A}/><w:ins w:id="5" ${A}/></w:p>`],
    ['an id that a number cannot hold exactly is Infinity', `<w:p><w:ins w:id="9007199254740993" ${A}/><w:ins w:id="2" ${A}/></w:p>`],
    ['the largest exact id', `<w:p><w:ins w:id="9007199254740991" ${A}/></w:p>`],
    ['same local names in another namespace are not revisions', `<w:p><x:ins x:id="99"/><x:del x:id="98"/><w:ins w:id="1" ${A}/></w:p>`],
    ['an unprefixed id attribute is not a Word id', `<w:p><w:ins id="77" ${A}/></w:p>`],
    ['customXml range markers', `<w:p><w:customXmlInsRangeStart w:id="21" ${A}/><w:customXmlInsRangeEnd w:id="21"/></w:p>`],
  ];

  for (const [label, body] of shapes) {
    it(`${label}: same count and same highest id as the old code`, () => {
      const document = parse(body);
      expect(countRevisionElements(document)).toBe(oldCount([document]));
      expect(countRevisions([document])).toBe(oldCount([document]));
      const expected = oldMaxRevisionId(document);
      expect(maxRevisionId(document)).toBe(expected);
    });
  }

  it('gives the figures the shapes are meant to show', () => {
    expect(maxRevisionId(parse(shapes[4][1]))).toBe(51);
    expect(countRevisionElements(parse(shapes[4][1]))).toBe(1);
    expect(maxRevisionId(parse(shapes[7][1]))).toBe(Infinity);
    expect(maxRevisionId(parse(shapes[9][1]))).toBe(1);
    expect(countRevisionElements(parse(shapes[9][1]))).toBe(1);
  });

  it('counts across several documents', () => {
    const a = parse(shapes[1][1]);
    const b = parse(shapes[3][1]);
    expect(countRevisions([a, b])).toBe(oldCount([a, b]));
    expect(countRevisions([])).toBe(0);
  });

  it('a very deep document does not overflow the stack and is still counted', () => {
    const depth = 20_000;
    const xml = `<w:document xmlns:w="${W}"><w:body>${'<w:sdt><w:sdtContent>'.repeat(depth)}<w:ins w:id="4" w:author="a" w:date="2026-01-01T00:00:00Z"/>${'</w:sdtContent></w:sdt>'.repeat(depth)}</w:body></w:document>`;
    const document = new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
    expect(countRevisionElements(document)).toBe(1);
    expect(maxRevisionId(document)).toBe(4);
  });

  it('visits an element once, and only elements of the names asked for', () => {
    const document = parse(`<w:p><w:r/><w:r/><w:t/></w:p>`);
    const seen: string[] = [];
    walkWordElements(document, new Set(['r']), el => seen.push(el.localName));
    expect(seen).toEqual(['r', 'r']);
  });

  it('agrees with the old code on every part of every .docx in the repository', () => {
    const files = execFileSync('git', ['ls-files', '*.docx'], { encoding: 'utf8' }).split('\n').filter(Boolean);
    let parts = 0;
    let withRevisions = 0;
    for (const file of files) {
      let zip: AdmZip;
      try { zip = new AdmZip(readFileSync(file)); } catch { continue; }
      for (const entry of zip.getEntries()) {
        if (!/^word\/[^/]+\.xml$/.test(entry.entryName)) continue;
        let document: Document;
        try { document = new DOMParser({ onError: () => undefined }).parseFromString(entry.getData().toString('utf8'), 'text/xml') as unknown as Document; } catch { continue; }
        if (!document.documentElement) continue;
        expect(countRevisionElements(document), `${file} ${entry.entryName}`).toBe(oldCount([document]));
        expect(maxRevisionId(document), `${file} ${entry.entryName}`).toBe(oldMaxRevisionId(document));
        parts++;
        if (oldCount([document]) > 0) withRevisions++;
      }
    }
    expect(parts).toBeGreaterThan(100);
    expect(withRevisions).toBeGreaterThanOrEqual(3);
  });
});
