/**
 * Allowlist normalizer proof: verified alias, default infill,
 * sequence sort, fail-closed gates, fixpoint, and the save-path hook.
 *
 * Every markup assertion reads NORMALIZED output (never assumptions about
 * the live DOM). Full-package assertions go through validateDocxBuffer.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer, executeInsert } from '@attentionform/astra-engine';
import {
  normalizePackage,
  previewNormalizations,
} from '../../packages/engine/src/ooxml/normalize';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const serializer = new XMLSerializer();

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'text/xml');
}

function textsOf(xml: string): string[] {
  const out: string[] = [];
  const nodes = parseXml(xml).getElementsByTagName('w:t');
  for (let i = 0; i < nodes.length; i++) out.push(nodes[i].textContent ?? '');
  return out;
}

function countTag(xml: string, tag: string): number {
  return parseXml(xml).getElementsByTagName(tag).length;
}

/** Minimal legacy part: alias + missing default + misordered container. */
const BROKEN_PART = [
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  '<w:body><w:tbl><w:tr><w:tc><w:tcPr>',
  '<w:shd w:fill="auto"/>',
  '<w:tcBdr><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBdr>',
  '<w:tcMar><w:top w:w="115" w:type="dxa"/></w:tcMar>',
  '</w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>',
].join('');

function tcPrChildTags(xml: string): string[] {
  const doc = parseXml(xml);
  const tcPr = doc.getElementsByTagName('w:tcPr')[0] as unknown as Element;
  const out: string[] = [];
  for (let i = 0; i < tcPr.childNodes.length; i++) {
    const c = tcPr.childNodes[i];
    if (c.nodeType === 1) out.push((c as Element).tagName);
  }
  return out;
}

/** Blank doc + one 3-row x 2-col table (6 cells), saved VALID. */
async function freshTableDoc(): Promise<{ buf: Buffer; xml: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']],
  } as any, ast, parser);
  const buf = parser.rezip(ast);
  return { buf, xml: new AdmZip(buf).readAsText('word/document.xml') };
}

function wEl(doc: Document, tag: string): Element {
  return doc.createElementNS(WORD_NS, tag);
}

/** Injects a defective tcPr (alias + val-less shd + wrong order) into every cell. */
function makeBrokenXml(cleanXml: string): { xml: string; cells: number } {
  const doc = parseXml(cleanXml);
  const tcs = doc.getElementsByTagName('w:tc');
  const cells = tcs.length;
  expect(cells).toBeGreaterThan(0);
  for (let i = 0; i < tcs.length; i++) {
    const tc = tcs[i] as unknown as Element;
    const tcPr = wEl(doc, 'w:tcPr');
    const shd = wEl(doc, 'w:shd');
    shd.setAttribute('w:fill', 'auto');
    const alias = wEl(doc, 'w:tcBdr');
    for (const side of ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV']) {
      const b = wEl(doc, side);
      b.setAttribute('w:val', 'single');
      b.setAttribute('w:sz', '4');
      b.setAttribute('w:space', '0');
      b.setAttribute('w:color', 'auto');
      alias.appendChild(b);
    }
    const mar = wEl(doc, 'w:tcMar');
    const marTop = wEl(doc, 'w:top');
    marTop.setAttribute('w:w', '115');
    marTop.setAttribute('w:type', 'dxa');
    mar.appendChild(marTop);
    tcPr.appendChild(shd);
    tcPr.appendChild(alias);
    tcPr.appendChild(mar);
    tc.insertBefore(tcPr, tc.firstChild);
  }
  return { xml: serializer.serializeToString(doc), cells };
}

function withDocumentXml(buf: Buffer, xml: string): Buffer {
  const zip = new AdmZip(buf);
  zip.updateFile('word/document.xml', Buffer.from(xml, 'utf8'));
  return zip.toBuffer();
}

describe('allowlist normalizer', () => {
  it('enumerates the three verified transforms with per-item locations', () => {
    const { parts, normalized } = normalizePackage({ 'word/document.xml': BROKEN_PART });
    expect(normalized).toEqual([
      'Renamed w:tcBdr \u2192 w:tcBorders at /w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:tcPr[1]/w:tcBorders[1]',
      'Infilled w:val="clear" on w:shd at /w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:tcPr[1]/w:shd[1]',
      'Reordered w:tcPr into schema order at /w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:tcPr[1]: [w:shd,w:tcBorders,w:tcMar] \u2192 [w:tcBorders,w:shd,w:tcMar]',
    ]);
    const out = parts['word/document.xml'];
    expect(countTag(out, 'w:tcBdr')).toBe(0);
    expect(countTag(out, 'w:tcBorders')).toBe(1);
    expect(tcPrChildTags(out)).toEqual(['w:tcBorders', 'w:shd', 'w:tcMar']);
    const shd = parseXml(out).getElementsByTagName('w:shd')[0] as unknown as Element;
    expect(shd.getAttribute('w:val')).toBe('clear');
    expect(shd.getAttribute('w:fill')).toBe('auto');
    expect(textsOf(out)).toEqual(['cell']);
  });

  it('leaves clean parts alone: empty list plus byte-identical output', async () => {
    const { xml } = await freshTableDoc();
    const res = normalizePackage({ 'word/document.xml': xml });
    expect(res.normalized).toEqual([]);
    expect(res.parts['word/document.xml'] === xml).toBe(true);
    expect(previewNormalizations({ 'word/document.xml': xml })).toEqual([]);
  });

  it('keeps unknown tags verbatim and sorts them last, stable', () => {
    const part = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:p><w:pPr><w:zzzCustom/><w:zzzSecond/><w:jc w:val="center"/></w:pPr>',
      '<w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>',
    ].join('');
    const { parts, normalized } = normalizePackage({ 'word/document.xml': part });
    const out = parts['word/document.xml'];
    expect(out).toContain('<w:zzzCustom/>');
    expect(out).toContain('<w:zzzSecond/>');
    const pPr = parseXml(out).getElementsByTagName('w:pPr')[0] as unknown as Element;
    const tags: string[] = [];
    for (let i = 0; i < pPr.childNodes.length; i++) {
      const c = pPr.childNodes[i];
      if (c.nodeType === 1) tags.push((c as Element).tagName);
    }
    expect(tags).toEqual(['w:jc', 'w:zzzCustom', 'w:zzzSecond']);
    expect(textsOf(out)).toEqual(['hi']);
    expect(normalized).toEqual([
      'Reordered w:pPr into schema order at /w:document[1]/w:body[1]/w:p[1]/w:pPr[1]: [w:zzzCustom,w:zzzSecond,w:jc] \u2192 [w:jc,w:zzzCustom,w:zzzSecond]',
    ]);
  });

  it('never adopts unverified aliases: tblBdr and pBdr stay untouched', () => {
    const part = [
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      '<w:body><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblBdr/></w:tblPr>',
      '<w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc>',
      '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:pBdr/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>',
      '</w:tc></w:tr></w:tbl></w:body></w:document>',
    ].join('');
    const { parts, normalized } = normalizePackage({ 'word/document.xml': part });
    expect(parts['word/document.xml']).toContain('w:tblBdr');
    expect(parts['word/document.xml']).toContain('w:pBdr');
    expect(normalized.some((line) => line.startsWith('Renamed'))).toBe(false);
    expect(textsOf(parts['word/document.xml'])).toEqual(['x']);
  });

  it('fails closed on syntax errors, in both write and preview paths', () => {
    const bad = { 'word/document.xml': '<w:document><w:body><w:p><w:r><w:t>oops' };
    expect(() => normalizePackage(bad)).toThrow(/E_VALIDATION/);
    expect(() => previewNormalizations(bad)).toThrow(/E_VALIDATION/);
    try {
      normalizePackage(bad);
      expect.unreachable();
    } catch (err: any) {
      expect(err.code).toBe('E_VALIDATION');
    }
  });

  it('is a fixpoint: second run is empty and byte-identical', () => {
    const first = normalizePackage({ 'word/document.xml': BROKEN_PART });
    expect(first.normalized.length).toBeGreaterThan(0);
    const second = normalizePackage(first.parts);
    expect(second.normalized).toEqual([]);
    expect(second.parts['word/document.xml'] === first.parts['word/document.xml']).toBe(true);
  });

  it('previews the same list as the write path without writing', () => {
    const input = { 'word/document.xml': BROKEN_PART };
    const before = JSON.stringify(input);
    const preview = previewNormalizations(input);
    const { normalized } = normalizePackage(input);
    expect(preview).toEqual(normalized);
    expect(JSON.stringify(input)).toBe(before);
    expect(preview.length).toBe(3);
  });

  it('heals a legacy package to VALID with content intact', async () => {
    const { buf, xml } = await freshTableDoc();
    const baseline = await validateDocxBuffer(buf);
    expect(baseline.ok).toBe(true);

    const { xml: brokenXml, cells } = makeBrokenXml(xml);
    expect(cells).toBe(6);
    const brokenBuf = withDocumentXml(buf, brokenXml);
    const broken = await validateDocxBuffer(brokenBuf);
    expect(broken.ok).toBe(false);

    const { parts, normalized } = normalizePackage({ 'word/document.xml': brokenXml });
    const tcPath = (r: number, c: number): string =>
      `/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${r}]/w:tc[${c}]`;
    const expected: string[] = [];
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expected.push(`Renamed w:tcBdr \u2192 w:tcBorders at ${tcPath(r, c)}/w:tcPr[1]/w:tcBorders[1]`);
      }
    }
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expected.push(`Infilled w:val="clear" on w:shd at ${tcPath(r, c)}/w:tcPr[1]/w:shd[1]`);
      }
    }
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expected.push(
          `Reordered w:tcPr into schema order at ${tcPath(r, c)}/w:tcPr[1]: [w:shd,w:tcBorders,w:tcMar] \u2192 [w:tcBorders,w:shd,w:tcMar]`
        );
      }
    }
    expect(normalized).toEqual(expected);
    expect(textsOf(parts['word/document.xml'])).toEqual(textsOf(brokenXml));

    const fixed = await validateDocxBuffer(withDocumentXml(buf, parts['word/document.xml']));
    expect(fixed.ok).toBe(true);
  });

  it('heals out-of-sequence tcBorders and pBdr children to VALID', async () => {
    const { buf, xml } = await freshTableDoc();
    const doc = parseXml(xml);
    const borderChild = (tag: string): Element => {
      const b = wEl(doc, tag);
      b.setAttribute('w:val', 'single');
      b.setAttribute('w:sz', '4');
      b.setAttribute('w:space', '0');
      b.setAttribute('w:color', 'auto');
      return b;
    };
    // First cell: fresh tcPr holding only a reversed tcBorders.
    const tc = doc.getElementsByTagName('w:tc')[0] as unknown as Element;
    const tcPr = wEl(doc, 'w:tcPr');
    const tcBorders = wEl(doc, 'w:tcBorders');
    for (const side of ['w:insideV', 'w:insideH', 'w:right', 'w:bottom', 'w:left', 'w:top']) {
      tcBorders.appendChild(borderChild(side));
    }
    tcPr.appendChild(tcBorders);
    tc.insertBefore(tcPr, tc.firstChild);
    // A paragraph without direct formatting: fresh pPr holding only a reversed pBdr.
    const paras = doc.getElementsByTagName('w:p');
    let target: Element | null = null;
    for (let i = 0; i < paras.length; i++) {
      const cand = paras[i] as unknown as Element;
      if (cand.getElementsByTagName('w:pPr').length === 0) {
        target = cand;
        break;
      }
    }
    expect(target).not.toBeNull();
    const pPr = wEl(doc, 'w:pPr');
    const pBdr = wEl(doc, 'w:pBdr');
    for (const side of ['w:bar', 'w:between', 'w:right', 'w:bottom', 'w:left', 'w:top']) {
      pBdr.appendChild(borderChild(side));
    }
    pPr.appendChild(pBdr);
    target!.insertBefore(pPr, target!.firstChild);
    const brokenXml = serializer.serializeToString(doc);

    const broken = await validateDocxBuffer(withDocumentXml(buf, brokenXml));
    expect(broken.ok).toBe(false);

    const { parts, normalized } = normalizePackage({ 'word/document.xml': brokenXml });
    expect(normalized).toEqual([
      'Reordered w:tcBorders into schema order at /w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:tcPr[1]/w:tcBorders[1]: [w:insideV,w:insideH,w:right,w:bottom,w:left,w:top] \u2192 [w:top,w:left,w:bottom,w:right,w:insideH,w:insideV]',
      'Reordered w:pBdr into schema order at /w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]/w:pPr[1]/w:pBdr[1] (node 03062A48): [w:bar,w:between,w:right,w:bottom,w:left,w:top] \u2192 [w:top,w:left,w:bottom,w:right,w:between,w:bar]',
    ]);
    const fixedDoc = parseXml(parts['word/document.xml']);
    const borderTags = (tag: string): string[] => {
      const el = fixedDoc.getElementsByTagName(tag)[0] as unknown as Element;
      const out: string[] = [];
      for (let i = 0; i < el.childNodes.length; i++) {
        const c = el.childNodes[i];
        if (c.nodeType === 1) out.push((c as Element).tagName);
      }
      return out;
    };
    expect(borderTags('w:tcBorders')).toEqual(['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV']);
    expect(borderTags('w:pBdr')).toEqual(['w:top', 'w:left', 'w:bottom', 'w:right', 'w:between', 'w:bar']);
    expect(textsOf(parts['word/document.xml'])).toEqual(textsOf(brokenXml));

    const fixed = await validateDocxBuffer(withDocumentXml(buf, parts['word/document.xml']));
    expect(fixed.ok).toBe(true);
  });

  it('unknown tags stay fail-closed at the gate after normalization', async () => {
    const { buf, xml } = await freshTableDoc();
    const doc = parseXml(xml);
    const tc = doc.getElementsByTagName('w:tc')[0] as unknown as Element;
    const tcPr = wEl(doc, 'w:tcPr');
    tcPr.appendChild(wEl(doc, 'w:zzzUnknown'));
    tc.insertBefore(tcPr, tc.firstChild);
    const strangeXml = serializer.serializeToString(doc);
    const { parts } = normalizePackage({ 'word/document.xml': strangeXml });
    expect(parts['word/document.xml']).toContain('w:zzzUnknown');
    const res = await validateDocxBuffer(withDocumentXml(buf, parts['word/document.xml']));
    expect(res.ok).toBe(false);
  });

  it('save-path hook heals on rezip, logs, then lands on an empty fixpoint', async () => {
    const { buf, xml } = await freshTableDoc();
    const { xml: brokenXml, cells } = makeBrokenXml(xml);
    expect(cells).toBe(6);
    const parser = new DocxParser(withDocumentXml(buf, brokenXml));

    const healed = parser.rezip();
    const tcPath = (r: number, c: number): string =>
      `/w:document[1]/w:body[1]/w:tbl[1]/w:tr[${r}]/w:tc[${c}]`;
    const expected: string[] = [];
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expected.push(`Renamed w:tcBdr \u2192 w:tcBorders at ${tcPath(r, c)}/w:tcPr[1]/w:tcBorders[1] (node t_001)`);
      }
    }
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expected.push(`Infilled w:val="clear" on w:shd at ${tcPath(r, c)}/w:tcPr[1]/w:shd[1] (node t_001)`);
      }
    }
    for (let r = 1; r <= 3; r++) {
      for (let c = 1; c <= 2; c++) {
        expected.push(
          `Reordered w:tcPr into schema order at ${tcPath(r, c)}/w:tcPr[1] (node t_001): [w:shd,w:tcBorders,w:tcMar] \u2192 [w:tcBorders,w:shd,w:tcMar]`
        );
      }
    }
    expect(parser.lastNormalizations).toEqual(expected);
    expect((await validateDocxBuffer(healed)).ok).toBe(true);
    expect(textsOf(new AdmZip(healed).readAsText('word/document.xml'))).toEqual(textsOf(xml));

    const healedAgain = parser.rezip();
    expect(parser.lastNormalizations).toEqual([]);
    expect(Buffer.compare(healedAgain, healed)).toBe(0);
  });

  it('save-path hook is a no-op on clean documents: empty log, identical bytes', async () => {
    const { buf } = await freshTableDoc();
    const parser = new DocxParser(buf);
    const first = parser.rezip();
    expect(parser.lastNormalizations).toEqual([]);
    const second = parser.rezip();
    expect(parser.lastNormalizations).toEqual([]);
    expect(Buffer.compare(second, first)).toBe(0);
  });
});
