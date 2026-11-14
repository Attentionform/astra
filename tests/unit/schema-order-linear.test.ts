/**
 * Canonicalisation stays linear in the size of the document, leaves a document that is already in
 * schema order untouched, and no loop in the engine edits the document while it reads a live
 * xmldom list unless it is on the reviewed list below.
 */
import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { DOMParser } from '@xmldom/xmldom';
import {
  canonicalizePartXml,
  sortChildrenBySchema,
  RPR_CHILD_ORDER,
} from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Document with `paragraphs` paragraphs of 20 formatted runs; the run properties of every other run are out of order. */
function makeDocument(paragraphs: number): string {
  let body = '';
  for (let p = 0; p < paragraphs; p++) {
    let runs = '';
    for (let r = 0; r < 20; r++) {
      const props = r % 2 === 0 ? '<w:sz w:val="22"/><w:b/>' : '<w:b/><w:sz w:val="22"/>';
      runs += `<w:r><w:rPr>${props}</w:rPr><w:t>run ${r}</w:t></w:r>`;
    }
    body += `<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0"/></w:pPr>${runs}</w:p>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;
}

/** The fastest of several runs: a pause from garbage collection or a busy runner only ever adds time to a run. */
function fastestMs(fn: () => void, times = 7): number {
  let best = Infinity;
  for (let i = 0; i < times; i++) {
    const t0 = process.hrtime.bigint();
    fn();
    best = Math.min(best, Number(process.hrtime.bigint() - t0) / 1e6);
  }
  return best;
}

describe('canonicalisation cost grows linearly', () => {
  it('doubling the document multiplies the time by less than 3 (a quadratic loop gives about 4)', () => {
    const small = makeDocument(150);
    const large = makeDocument(300);
    canonicalizePartXml(makeDocument(20)); // warm up
    // A shared CI runner can slow either measurement for a while, so the test measures again up to five times
    // and passes on the first ratio under the limit. A quadratic loop gives about 4 every time; local runs give about 2.
    const ratios: number[] = [];
    for (let attempt = 0; attempt < 5; attempt++) {
      const t1 = fastestMs(() => canonicalizePartXml(small));
      const t2 = fastestMs(() => canonicalizePartXml(large));
      ratios.push(t2 / t1);
      if (t2 / t1 < 3) return;
    }
    expect(Math.min(...ratios)).toBeLessThan(3);
  });

  it('sorts out-of-order run properties and keeps in-order input as it is', () => {
    const out = canonicalizePartXml(makeDocument(1));
    expect(out).toContain('<w:rPr><w:b/><w:sz w:val="22"/></w:rPr>');
    expect(out).not.toContain('<w:rPr><w:sz w:val="22"/><w:b/></w:rPr>');
    const again = canonicalizePartXml(out);
    expect(again).toBe(out);
  });
});

describe('sortChildrenBySchema', () => {
  const parse = (xml: string): Element =>
    new DOMParser().parseFromString(`<w:rPr xmlns:w="${W}">${xml}</w:rPr>`, 'text/xml').documentElement as Element;

  it('does not touch the children when they are already in order', () => {
    const el = parse('<w:b/><w:x/><w:sz w:val="22"/>');
    let moves = 0;
    const remove = el.removeChild.bind(el);
    const append = el.appendChild.bind(el);
    el.removeChild = ((c: Node) => { moves++; return remove(c); }) as typeof el.removeChild;
    el.appendChild = ((c: Node) => { moves++; return append(c); }) as typeof el.appendChild;
    sortChildrenBySchema(el, RPR_CHILD_ORDER);
    expect(moves).toBe(0);
  });

  it('sorts listed children and keeps an unlisted element in its slot', () => {
    const el = parse('<w:sz w:val="22"/><w:x/><w:b/>');
    sortChildrenBySchema(el, RPR_CHILD_ORDER);
    const names: string[] = [];
    for (let c = el.firstChild; c; c = c.nextSibling) names.push((c as Element).tagName);
    expect(names).toEqual(['w:b', 'w:x', 'w:sz']);
  });
});

/**
 * Loops of the form `for (let i = 0; i < list.length; i++)` over a live `getElementsByTagName` list whose
 * body edits the document. Each one was read; the body never adds or removes an element of the iterated tag,
 * and none is on the edit or save path that was measured, so none is changed. A new loop of this shape must
 * either use a snapshot or be added here with a reason.
 */
const REVIEWED: Record<string, number> = {
  'ooxml/typography.ts': 3,
  'operations/clearFormat.ts': 2,
  'operations/image.ts': 1,
  'operations/table.ts': 1,
  'parser.ts': 2,
  // Read-only walk: children are pushed on a separate work list and nothing is changed while iterating.
  'visual/references.ts': 1,
};
const MUTATION = /\.(setAttribute|setAttributeNS|removeAttribute|removeChild|appendChild|insertBefore|replaceChild|remove|setAttributeNode)\(|\.textContent\s*=[^=]|\.nodeValue\s*=[^=]|\.data\s*=[^=]/;

function listFiles(dir: string, out: string[] = []): string[] {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) listFiles(p, out);
    else if (p.endsWith('.ts') && !p.endsWith('.d.ts')) out.push(p);
  }
  return out;
}

function bodyOf(src: string, open: number): string {
  let depth = 0;
  for (let i = open; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') {
      depth--;
      if (depth === 0) return src.slice(open, i + 1);
    }
  }
  return '';
}

describe('live-list guard', () => {
  it('finds no unreviewed loop that edits the document while reading a live list', () => {
    const root = path.resolve(__dirname, '../../packages/engine/src');
    const found: Record<string, number> = {};
    for (const file of listFiles(root)) {
      const src = fs.readFileSync(file, 'utf8');
      const live = new Set<string>();
      const decl = /(?:const|let)\s+(\w+)\s*=\s*([^;\n]*getElementsByTagName(?:NS)?\([^;\n]*)/g;
      let m: RegExpExecArray | null;
      while ((m = decl.exec(src))) {
        if (/Array\.from|\[\.\.\.|\.slice\(|\.filter\(|\.map\(|snapshot\(/.test(m[2])) continue;
        if (/\)\s*\[\d+\]|\)\s*\?\.\s*\[/.test(m[2])) continue;
        live.add(m[1]);
      }
      const loop = /for\s*\(\s*(?:let|var)\s+(\w+)\s*=\s*0\s*;\s*\1\s*<\s*([\w.]+?)\.length\s*;[^)]*\)\s*\{/g;
      while ((m = loop.exec(src))) {
        if (!live.has(m[2]) && !/getElementsByTagName/.test(m[2])) continue;
        const body = bodyOf(src, m.index + m[0].length - 1);
        if (!MUTATION.test(body)) continue;
        const rel = path.relative(root, file).split(path.sep).join('/');
        found[rel] = (found[rel] ?? 0) + 1;
      }
    }
    expect(found).toEqual(REVIEWED);
  });
});
