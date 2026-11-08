import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { createBlankDocxBuffer, sanitizeDocxBuffer, sanitizeDocumentXml, validateDocxBuffer } from '@attentionform/astra-engine';
import { resolveRevisions } from '../../packages/engine/src/ooxml/revisions';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STAMP = 'w:author="Alice" w:date="2026-01-01T00:00:00Z"';

function docWith(body: string, extra?: (zip: AdmZip) => void): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  extra?.(zip);
  return zip.toBuffer();
}

const sanitize = (buffer: Buffer) => sanitizeDocxBuffer(buffer, { acceptAllRevisions: true } as never);
const out = (result: unknown): Buffer => {
  const r = result as { buffer?: Buffer; output?: Buffer };
  return Buffer.isBuffer(result) ? result : (r.buffer ?? r.output)!;
};
const text = (buffer: Buffer, part = 'word/document.xml') => new AdmZip(buffer).readAsText(part);

describe('sanitize with acceptAllRevisions fails closed (FX-02, FX-03)', () => {
  const unresolvable: Record<string, { body: string; kind: string }> = {
    'a deleted paragraph mark next to a section break': {
      kind: 'del',
      body: `<w:p><w:pPr><w:rPr><w:del w:id="1" ${STAMP}/></w:rPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>`,
    },
    'an orphan moveFromRangeStart': { kind: 'moveFromRangeStart', body: `<w:p><w:moveFromRangeStart w:id="5" w:name="m1" ${STAMP}/><w:r><w:t>keep</w:t></w:r></w:p>` },
    'an orphan moveToRangeEnd': { kind: 'moveToRangeEnd', body: `<w:p><w:r><w:t>keep</w:t></w:r><w:moveToRangeEnd w:id="6"/></w:p>` },
    'an incomplete moveFrom with text': { kind: 'moveFrom', body: `<w:p><w:moveFrom w:id="1" ${STAMP}><w:r><w:t>SECRETMOVED</w:t></w:r></w:moveFrom><w:r><w:t>keep</w:t></w:r></w:p>` },
    'customXml markers around deleted content': {
      kind: 'customXmlDelRangeStart',
      body: `<w:p><w:customXmlDelRangeStart w:id="8" ${STAMP}/><w:del w:id="9" ${STAMP}><w:r><w:delText>gone</w:delText></w:r></w:del><w:customXmlDelRangeEnd w:id="8"/></w:p>`,
    },
  };

  for (const [name, { body, kind }] of Object.entries(unresolvable)) {
    it(`refuses ${name} and returns no file`, async () => {
      let thrown: Error | undefined;
      try { await sanitize(docWith(body)); } catch (error) { thrown = error as Error; }
      expect(thrown, 'sanitize must throw').toBeDefined();
      expect(thrown!.message).toMatch(/E_VALIDATION/);
      expect(thrown!.message).toContain('word/document.xml');
      expect(thrown!.message).toContain(kind);
      expect(thrown!.message).toMatch(/without acceptAllRevisions/);
    });
  }

  it('refuses an unresolvable revision in a header and names the part', async () => {
    const hdr = `<w:hdr xmlns:w="${W}"><w:p><w:moveFrom w:id="1" ${STAMP}><w:r><w:t>HDRSECRET</w:t></w:r></w:moveFrom></w:p></w:hdr>`;
    const buffer = docWith('<w:p><w:r><w:t>body</w:t></w:r></w:p>', zip => {
      zip.addFile('word/header1.xml', Buffer.from(hdr));
      zip.updateFile('[Content_Types].xml', Buffer.from(zip.readAsText('[Content_Types].xml').replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')));
    });
    expect(() => sanitize(buffer)).toThrow(/word\/header1\.xml/);
  });

  it('removes every revision and all reviewer metadata from a resolvable document', async () => {
    const body = `<w:p><w:ins w:id="1" ${STAMP}><w:r><w:t>kept</w:t></w:r></w:ins><w:del w:id="2" ${STAMP}><w:r><w:delText>GONE</w:delText></w:r></w:del></w:p>` +
      `<w:p><w:customXmlInsRangeStart w:id="3" ${STAMP}/><w:r><w:t>inserted block</w:t></w:r><w:customXmlInsRangeEnd w:id="3"/></w:p>` +
      `<w:p><w:pPr><w:pPrChange w:id="4" ${STAMP}><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>formatted</w:t></w:r></w:p>`;
    const buffer = out(await sanitize(docWith(body)));
    const xml = text(buffer);
    expect(xml).not.toContain('Alice');
    expect(xml).not.toContain('GONE');
    expect(xml).not.toMatch(/<w:(ins|del|customXml\w+|pPrChange)[ >/]/);
    expect(xml).toContain('kept');
    expect(xml).toContain('inserted block');
    expect(xml).toContain('formatted');
    expect((await validateDocxBuffer(buffer)).ok).toBe(true);
  });

  it('a document with no revisions is unchanged by the revision pass', async () => {
    const buffer = out(await sanitize(docWith('<w:p><w:r><w:t>plain</w:t></w:r></w:p>')));
    expect(text(buffer)).toContain('plain');
    expect((await validateDocxBuffer(buffer)).ok).toBe(true);
  });

  it('sanitizeDocumentXml refuses unpaired moves instead of returning them', () => {
    const xml = `<w:p xmlns:w="${W}"><w:moveFrom w:id="1" ${STAMP}><w:r><w:t>SECRET</w:t></w:r></w:moveFrom></w:p>`;
    expect(() => sanitizeDocumentXml(xml, { acceptAllRevisions: true })).toThrow(/E_VALIDATION.*moveFrom/s);
  });
});

describe('the resolver handles the safe customXml marker cases (FX-03, decision D-3)', () => {
  const resolve = (body: string, mode: 'accept' | 'reject') => {
    const doc = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${body}</w:body>`, 'text/xml');
    const result = resolveRevisions([doc], mode);
    return { result, xml: new XMLSerializer().serializeToString(doc) };
  };
  const insPair = `<w:customXmlInsRangeStart w:id="1" ${STAMP}/><w:p><w:r><w:t>TOC</w:t></w:r></w:p><w:customXmlInsRangeEnd w:id="1"/>`;

  it('accept drops insert markers and keeps the content', () => {
    const { result, xml } = resolve(insPair, 'accept');
    expect(result.unresolved).toEqual([]);
    expect(xml).toContain('TOC');
    expect(xml).not.toContain('customXml');
  });

  it('accept drops move-to markers and keeps the content', () => {
    const { result, xml } = resolve(`<w:customXmlMoveToRangeStart w:id="2" ${STAMP}/><w:p><w:r><w:t>moved</w:t></w:r></w:p><w:customXmlMoveToRangeEnd w:id="2"/>`, 'accept');
    expect(result.unresolved).toEqual([]);
    expect(xml).toContain('moved');
    expect(xml).not.toContain('customXml');
  });

  it('reject removes the inserted content between the markers', () => {
    const { result, xml } = resolve(`<w:p><w:r><w:t>before</w:t></w:r></w:p>${insPair}<w:p><w:r><w:t>after</w:t></w:r></w:p>`, 'reject');
    expect(result.unresolved).toEqual([]);
    expect(xml).toContain('before');
    expect(xml).toContain('after');
    expect(xml).not.toContain('TOC');
    expect(xml).not.toContain('customXml');
  });

  it('reject with an end marker that has no start stays unresolved', () => {
    const { result, xml } = resolve(`<w:p><w:r><w:t>keep</w:t></w:r></w:p><w:customXmlInsRangeEnd w:id="1"/>`, 'reject');
    expect(result.unresolved.map(u => u.kind)).toEqual(['customXmlInsRangeEnd']);
    expect(xml).toContain('keep');
  });

  it('delete markers around content stay unresolved in both modes', () => {
    const body = `<w:customXmlDelRangeStart w:id="3" ${STAMP}/><w:p><w:r><w:t>x</w:t></w:r></w:p><w:customXmlDelRangeEnd w:id="3"/>`;
    for (const mode of ['accept', 'reject'] as const) {
      const { result } = resolve(body, mode);
      expect(result.unresolved.length).toBeGreaterThan(0);
    }
  });

  it('delete markers that enclose nothing are dropped', () => {
    const { result, xml } = resolve(`<w:customXmlDelRangeStart w:id="3" ${STAMP}/><w:customXmlDelRangeEnd w:id="3"/><w:p><w:r><w:t>x</w:t></w:r></w:p>`, 'accept');
    expect(result.unresolved).toEqual([]);
    expect(xml).not.toContain('customXml');
  });
});
