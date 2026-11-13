import { describe, it, expect } from 'vitest';
import {
  ORDERED_CONTAINERS,
  reorderPartXml,
  verifyOrderOnlyRepair,
} from '../../packages/engine/src/ooxml/normalize';

/**
 * `verifyOrderOnlyRepair` is the proof that a repaired part differs from its original only in the order of
 * listed children inside the ordered containers. It reads a fixed allow-list, never the list a repair reports,
 * and it does not use the sort code, so each rejected case below builds the "repaired" text by hand.
 */
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS = `xmlns:w="${W}"`;
const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const doc = (body: string, prolog = PROLOG, epilog = ''): string =>
  `${prolog}<w:document ${NS}><w:body>${body}</w:body></w:document>${epilog}`;
const verify = (a: string, b: string): boolean => verifyOrderOnlyRepair('word/document.xml', a, b);

const RUN_BAD = '<w:r><w:rPr><w:rFonts w:ascii="A"/><w:sz w:val="28"/><w:color w:val="000000"/></w:rPr><w:t>One</w:t></w:r>';
const RUN_OK = '<w:r><w:rPr><w:rFonts w:ascii="A"/><w:color w:val="000000"/><w:sz w:val="28"/></w:rPr><w:t>One</w:t></w:r>';
const P2 = '<w:p><w:r><w:t>Two</w:t></w:r></w:p>';
const P1 = `<w:p>${RUN_BAD}</w:p>`;

const repaired = (original: string) => reorderPartXml('word/document.xml', original);

describe('reorderPartXml', () => {
  it('reorders only, and reports the container xPath', () => {
    const r = repaired(doc(P1 + P2));
    expect(r.reordered).toHaveLength(1);
    expect(r.containers).toEqual(['/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:rPr[1]']);
    expect(r.xml).toContain('<w:color w:val="000000"/><w:sz w:val="28"/>');
  });

  it('returns the exact input when nothing needs reordering', () => {
    const clean = doc(`<w:p>${RUN_OK}</w:p>`);
    const r = repaired(clean);
    expect(r.reordered).toEqual([]);
    expect(r.xml).toBe(clean);
  });

  it('does not rename or infill: an alias tag and a shd without val stay as they are', () => {
    const odd = doc('<w:p><w:r><w:rPr><w:shd w:fill="FF0000"/></w:rPr><w:t>x</w:t></w:r></w:p><w:tcBdr/>');
    expect(repaired(odd).xml).toBe(odd);
  });
});

describe('the allow-list', () => {
  it('holds exactly the containers the normalizer sorts, and none whose child order carries meaning', () => {
    expect([...ORDERED_CONTAINERS.keys()].sort()).toEqual(['pBdr', 'pPr', 'rPr', 'tblPr', 'tcBorders', 'tcPr']);
    for (const tag of ['body', 'p', 'tc', 'tr', 'tbl', 'sectPr', 'tblBorders']) expect(ORDERED_CONTAINERS.has(tag)).toBe(false);
  });
});

describe('verifyOrderOnlyRepair: accepted', () => {
  const original = doc(P1 + P2);
  const good = repaired(original).xml;

  it('the real repair', () => expect(verify(original, good)).toBe(true));

  it('a hand-built reorder inside the listed container', () => {
    expect(verify(original, doc(`<w:p>${RUN_OK}</w:p>${P2}`))).toBe(true);
  });

  it('does not read any list from the repair: it needs none', () => {
    expect(verifyOrderOnlyRepair.length).toBe(3);
  });

  it('nested repairs: a pPr out of order whose inner rPr and pBdr are also out of order', () => {
    const nested = doc(
      '<w:p><w:pPr><w:rPr><w:sz w:val="20"/><w:color w:val="111111"/></w:rPr><w:jc w:val="left"/><w:pBdr><w:bottom w:val="single"/><w:top w:val="single"/></w:pBdr><w:keepNext/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'
    );
    const r = repaired(nested);
    expect(r.containers).toHaveLength(3);
    expect(verify(nested, r.xml)).toBe(true);
    expect(r.xml.indexOf('<w:keepNext/>')).toBeLessThan(r.xml.indexOf('<w:pBdr>'));
    expect(r.xml.indexOf('<w:pBdr>')).toBeLessThan(r.xml.indexOf('<w:jc'));
  });

  it('a tcPr with an out-of-order tcBorders', () => {
    const t = doc(
      '<w:tbl><w:tr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single"/><w:top w:val="single"/></w:tcBorders><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'
    );
    const r = repaired(t);
    expect(r.reordered.length).toBeGreaterThanOrEqual(2);
    expect(verify(t, r.xml)).toBe(true);
  });

  it('the same container tag reordered in many places', () => {
    const many = doc(P1.repeat(25) + P2);
    const r = repaired(many);
    expect(r.containers).toHaveLength(25);
    expect(verify(many, r.xml)).toBe(true);
  });

  it('a unlisted element in an ordered container stays in its slot and is accepted', () => {
    const odd = doc('<w:p><w:pPr><w:jc w:val="left"/><w:ins w:id="1" w:author="a"/><w:keepNext/></w:pPr></w:p>');
    const r = repaired(odd);
    expect(r.reordered).toHaveLength(1);
    expect(verify(odd, r.xml)).toBe(true);
  });

  it('text, comments and processing instructions elsewhere are compared in sequence and accepted when equal', () => {
    const withMixed = doc(`<!-- keep --><?pi data?>${P1}<w:p><w:r><w:t>a</w:t></w:r></w:p>`);
    expect(verify(withMixed, repaired(withMixed).xml)).toBe(true);
  });

  it('a part with a BOM, standalone="yes", a prolog comment and a trailing comment keeps all of them', () => {
    const original2 = `\uFEFF${PROLOG}\n<!--c-->${doc(P1, '')}\n<!--e-->\n`;
    const r = repaired(original2);
    expect(r.xml.startsWith('\uFEFF<?xml')).toBe(true);
    expect(r.xml).toContain('standalone="yes"');
    expect(r.xml).toContain('<!--c-->');
    expect(r.xml).toContain('<!--e-->');
    expect(verify(original2, r.xml)).toBe(true);
  });

  it('a prefix other than w bound to the WordprocessingML namespace: nothing is sorted and the verifier accepts the unchanged text', () => {
    const aliased = `${PROLOG}<x:document xmlns:x="${W}"><x:body><x:p><x:r><x:rPr><x:sz x:val="28"/><x:color x:val="000000"/></x:rPr></x:r></x:p></x:body></x:document>`;
    const r = repaired(aliased);
    expect(r.reordered).toEqual([]);
    expect(verify(aliased, aliased)).toBe(true);
  });

  it('a child that carries its own namespace declaration is reordered with its container and keeps the declaration', () => {
    const local = doc(
      '<w:p><w:pPr><w:jc w:val="left"/><w:keepNext xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/></w:pPr></w:p>'
    );
    const r = repaired(local);
    expect(r.reordered).toHaveLength(1);
    expect(r.xml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
    expect(verify(local, r.xml)).toBe(true);
  });
});

describe('verifyOrderOnlyRepair: rejected', () => {
  const original = doc(P1 + P2);
  const good = repaired(original).xml;

  it('a changed attribute value in the same part', () => expect(verify(original, good.replace('w:val="28"', 'w:val="29"'))).toBe(false));
  it('changed text', () => expect(verify(original, good.replace('>One<', '>Uno<'))).toBe(false));
  it('a removed child', () => expect(verify(original, good.replace('<w:rFonts w:ascii="A"/>', ''))).toBe(false));
  it('a duplicated child', () => {
    expect(verify(original, good.replace('<w:sz w:val="28"/>', '<w:sz w:val="28"/><w:sz w:val="28"/>'))).toBe(false);
  });
  it('an added element', () => expect(verify(original, good.replace('</w:body>', '<w:p/></w:body>'))).toBe(false));

  it('two body paragraphs swapped (a container that is not in the allow-list)', () => {
    const before = doc(P2 + P1);
    expect(verify(before, doc(`<w:p>${RUN_OK}</w:p>${P2}`))).toBe(false);
  });

  it('the reorder is never excused by a list the caller supplies: swapping body children is refused regardless', () => {
    const before = doc('<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>');
    const after = doc('<w:p><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>A</w:t></w:r></w:p>');
    expect(verify(before, after)).toBe(false);
  });

  it('a swap inside a container that resembles an allow-list one (tblBorders)', () => {
    const before = doc('<w:tbl><w:tblPr><w:tblBorders><w:bottom w:val="single"/><w:top w:val="single"/></w:tblBorders></w:tblPr></w:tbl>');
    const after = doc('<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single"/><w:bottom w:val="single"/></w:tblBorders></w:tblPr></w:tbl>');
    expect(verify(before, after)).toBe(false);
  });

  it('duplicate-looking children where one is altered', () => {
    const before = doc('<w:p><w:r><w:rPr><w:b w:val="1"/><w:b w:val="0"/><w:color w:val="000000"/></w:rPr></w:r></w:p>');
    const after = doc('<w:p><w:r><w:rPr><w:b w:val="1"/><w:b w:val="1"/><w:color w:val="000000"/></w:rPr></w:r></w:p>');
    expect(verify(before, after)).toBe(false);
  });

  it('an unlisted element moved to another slot in an ordered container', () => {
    const before = doc('<w:p><w:pPr><w:keepNext/><w:ins w:id="1" w:author="a"/><w:jc w:val="left"/></w:pPr></w:p>');
    const after = doc('<w:p><w:pPr><w:ins w:id="1" w:author="a"/><w:keepNext/><w:jc w:val="left"/></w:pPr></w:p>');
    expect(verify(before, after)).toBe(false);
  });

  it('text, a comment or a processing instruction between the children of a reordered container (ambiguous)', () => {
    for (const filler of ['text', '<!--c-->', '<?pi x?>']) {
      const before = doc(`<w:p><w:pPr><w:jc w:val="left"/>${filler}<w:keepNext/></w:pPr></w:p>`);
      const after = doc(`<w:p><w:pPr><w:keepNext/>${filler}<w:jc w:val="left"/></w:pPr></w:p>`);
      expect(verify(before, after)).toBe(false);
    }
  });

  it('the same filler in a container that is not reordered is accepted', () => {
    const same = doc('<w:p><w:pPr><w:keepNext/>text<w:jc w:val="left"/></w:pPr></w:p>');
    expect(verify(same, same)).toBe(true);
  });

  it('a removed comment or processing instruction', () => {
    const withComment = doc(`<!-- keep -->${P1}${P2}`);
    expect(verify(withComment, repaired(withComment).xml.replace('<!-- keep -->', ''))).toBe(false);
  });

  it('a different XML declaration, prolog comment or epilog', () => {
    const a = doc(P1, `${PROLOG}<!--p-->`, '<!--e-->');
    const r = repaired(a).xml;
    expect(verify(a, r)).toBe(true);
    expect(verify(a, r.replace(' standalone="yes"', ''))).toBe(false);
    expect(verify(a, r.replace('<!--p-->', '<!--q-->'))).toBe(false);
    expect(verify(a, r.replace('<!--e-->', ''))).toBe(false);
  });

  it('namespace changes: a dropped, redundant or rebound declaration', () => {
    const local = doc('<w:p><w:pPr><w:jc w:val="left"/><w:keepNext xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/></w:pPr></w:p>');
    const r = repaired(local).xml;
    expect(verify(local, r.replace(' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', ''))).toBe(false);
    expect(verify(local, r.replace('<w:jc ', '<w:jc xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '))).toBe(false);
    expect(verify(local, r.replace('word/2010/wordml', 'word/2010/other'))).toBe(false);
  });

  it('fails closed on text that does not parse, and when equivalence cannot be shown', () => {
    expect(verify(original, '<w:document><w:body>')).toBe(false);
    const unbound = `${PROLOG}<w:document><w:body/></w:document>`;
    expect(verify(unbound, unbound)).toBe(false);
  });
});

describe('verifyOrderOnlyRepair: random documents', () => {
  const CHILDREN = ['<w:keepNext/>', '<w:keepLines/>', '<w:spacing w:before="1"/>', '<w:jc w:val="left"/>', '<w:outlineLvl w:val="1"/>'];
  function mulberry(seed: number) {
    return () => {
      seed |= 0;
      seed = (seed + 0x6d2b79f5) | 0;
      let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  it('shuffling listed children is accepted, and one random non-order mutation is rejected (200 documents)', () => {
    const rnd = mulberry(7);
    for (let n = 0; n < 200; n++) {
      const paras: string[] = [];
      const shuffled: string[] = [];
      for (let p = 0; p < 1 + Math.floor(rnd() * 4); p++) {
        const kids = CHILDREN.filter(() => rnd() < 0.7);
        const copy = [...kids].sort(() => rnd() - 0.5);
        paras.push(`<w:p><w:pPr>${kids.join('')}</w:pPr><w:r><w:t>t${p}</w:t></w:r></w:p>`);
        shuffled.push(`<w:p><w:pPr>${copy.join('')}</w:pPr><w:r><w:t>t${p}</w:t></w:r></w:p>`);
      }
      const a = doc(paras.join(''));
      const b = doc(shuffled.join(''));
      expect(verify(a, b)).toBe(true);
      const mutations = [b.replace('<w:t>t0</w:t>', '<w:t>tX</w:t>'), b.replace('</w:body>', '<w:p/></w:body>'), b.replace('<w:pPr>', '<w:pPr w:x="1">')];
      expect(verify(a, mutations[n % 3])).toBe(false);
    }
  });
});
