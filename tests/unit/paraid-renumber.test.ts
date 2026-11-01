import { describe, it, expect, afterEach } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, getDuplicateParaIdPolicy, setDuplicateParaIdPolicy } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

function pkg(body: string): Buffer {
  const zip = new AdmZip();
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
  zip.addFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}"><w:body>${body}</w:body></w:document>`,
      'utf8'
    )
  );
  return zip.toBuffer();
}
const para = (id: string, text: string) => `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const alt = (choice: string, fallback: string) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><txbx><w:txbxContent>${choice}</w:txbxContent></txbx></w:drawing></mc:Choice>` +
  `<mc:Fallback><w:pict><txbx><w:txbxContent>${fallback}</w:txbxContent></txbx></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

const doc = (buf: Buffer) => new AdmZip(buf).readAsText('word/document.xml');
const ids = (xml: string) => [...xml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1]);
const texts = (xml: string) => [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]);

afterEach(() => setDuplicateParaIdPolicy('refuse'));

describe('renumbering paragraphs that share a w14:paraId', () => {
  it('keeps the first id, gives the later copies new unique ids, keeps the text and logs each change', () => {
    const parser = new DocxParser(pkg(para('7B904439', 'one') + para('7B904439', 'two') + para('7B904439', 'three')));
    const out = parser.rezip({ duplicateParaIds: 'renumber', reproducible: true });
    const xml = doc(out);
    const found = ids(xml);
    expect(found).toHaveLength(3);
    expect(new Set(found).size).toBe(3);
    expect(found[0]).toBe('7B904439');
    expect(texts(xml)).toEqual(['one', 'two', 'three']);
    const renumbered = parser.lastNormalizations.filter((l) => l.startsWith('Renumbered duplicate'));
    expect(renumbered).toHaveLength(2);
    expect(renumbered[0]).toContain('"7B904439"');
    expect(renumbered[0]).toContain('/w:document[1]/w:body[1]/w:p[2]');
    expect(parser.lastParaIdRenumbering).toHaveLength(2);
  });

  it('a second save changes nothing and two reproducible saves give the same bytes', () => {
    const input = pkg(para('7B904439', 'one') + para('7B904439', 'two'));
    const first = new DocxParser(input).rezip({ duplicateParaIds: 'renumber', reproducible: true });
    const second = new DocxParser(input).rezip({ duplicateParaIds: 'renumber', reproducible: true });
    expect(first.equals(second)).toBe(true);
    const again = new DocxParser(first);
    expect(again.getDuplicateParaIds()).toEqual([]);
    const third = again.rezip({ duplicateParaIds: 'renumber', reproducible: true });
    expect(doc(third)).toBe(doc(first));
    expect(again.lastParaIdRenumbering).toEqual([]);
  });

  it('a repeating section of 15 copies ends with 15 distinct ids', () => {
    const copies = Array.from({ length: 15 }, () => para('2A2A2A2A', 'row')).join('');
    const parser = new DocxParser(pkg(copies));
    const xml = doc(parser.rezip({ duplicateParaIds: 'renumber', reproducible: true }));
    expect(new Set(ids(xml)).size).toBe(15);
    expect(texts(xml)).toEqual(Array(15).fill('row'));
  });

  it('leaves the two copies of a mirrored mc:Choice / mc:Fallback text box alone', () => {
    const inner = para('22222222', 'a') + para('33333333', 'b');
    const body = `<w:p w14:paraId="11111111">${alt(inner, inner)}</w:p>`;
    const parser = new DocxParser(pkg(body));
    const out = parser.rezip({ duplicateParaIds: 'renumber', reproducible: true });
    expect(ids(doc(out)).sort()).toEqual(['11111111', '22222222', '22222222', '33333333', '33333333']);
    expect(parser.lastParaIdRenumbering).toEqual([]);
  });

  it('renumbers a body paragraph that repeats the id of a mirrored text box and keeps the mirror', () => {
    const body = `<w:p w14:paraId="11111111">${alt(para('22222222', 'a'), para('22222222', 'a'))}</w:p>` + para('22222222', 'c');
    const parser = new DocxParser(pkg(body));
    const out = parser.rezip({ duplicateParaIds: 'renumber', reproducible: true });
    const found = ids(doc(out));
    expect(found.filter((i) => i === '22222222')).toHaveLength(2);
    expect(new Set(found).size).toBe(3);
    expect(parser.lastParaIdRenumbering).toHaveLength(1);
  });

  it('the refuse policy keeps today\'s error and changes nothing', () => {
    const parser = new DocxParser(pkg(para('7B904439', 'one') + para('7B904439', 'two')));
    expect(() => parser.rezip({ duplicateParaIds: 'refuse' })).toThrow(/Duplicate Guard/);
    expect(parser.getDuplicateParaIds()).toHaveLength(1);
  });

  it('the library default is refuse, and the process-wide policy can be switched', () => {
    expect(getDuplicateParaIdPolicy()).toBe('refuse');
    const input = pkg(para('7B904439', 'one') + para('7B904439', 'two'));
    expect(() => new DocxParser(input).rezip()).toThrow(/Duplicate Guard/);
    setDuplicateParaIdPolicy('renumber');
    expect(() => new DocxParser(input).rezip()).not.toThrow();
  });

  it('validateOOXML with the renumber policy no longer reports the duplicate', async () => {
    const parser = new DocxParser(pkg(para('7B904439', 'one') + para('7B904439', 'two')));
    const res = await parser.validateOOXML({ duplicateParaIds: 'renumber' });
    const dup = (res.errors as unknown as { errorType?: string }[]).filter((e) => e.errorType === 'DuplicateParaId');
    expect(dup).toEqual([]);
    expect(parser.getDuplicateParaIds()).toEqual([]);
  });

  it('the session id of a renumbered paragraph matches the saved id', () => {
    const parser = new DocxParser(pkg(para('7B904439', 'one') + para('7B904439', 'two')));
    expect(parser.getStructuralSummary().nodes[1].id).toBe('7B904439_dup');
    parser.rezip({ duplicateParaIds: 'renumber', reproducible: true });
    const second = parser.getStructuralSummary().nodes[1].id;
    expect(second).toMatch(/^[0-9A-F]{8}$/);
    expect(second).not.toBe('7B904439');
  });
});
