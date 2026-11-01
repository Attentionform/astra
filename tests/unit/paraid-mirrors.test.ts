import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

function pkg(body: string): Buffer {
  const zip = new AdmZip();
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
  zip.addFile('word/document.xml', Buffer.from(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}"><w:body>${body}</w:body></w:document>`, 'utf8'));
  return zip.toBuffer();
}
const textBox = (id: string, text: string) =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const alt = (choice: string, fallback: string) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><txbx><w:txbxContent>${choice}</w:txbxContent></txbx></w:drawing></mc:Choice>` +
  `<mc:Fallback><w:pict><txbx><w:txbxContent>${fallback}</w:txbxContent></txbx></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

describe('paragraph ids repeated by mc:Choice and mc:Fallback', () => {
  it('a text box written twice is not a duplicate and the document saves', () => {
    const body = `<w:p w14:paraId="11111111">${alt(textBox('22222222', 'a') + textBox('33333333', 'b'), textBox('22222222', 'a') + textBox('33333333', 'b'))}</w:p>`;
    const parser = new DocxParser(pkg(body));
    expect(parser.getDuplicateParaIds()).toEqual([]);
    expect(() => parser.rezip()).not.toThrow();
  });

  it('a text box nested in a text box is still not a duplicate', () => {
    const inner = (t: string) => alt(textBox('44444444', t), textBox('44444444', t));
    const outerParagraph = `<w:p w14:paraId="22222222">${inner('x')}</w:p>`;
    const body = `<w:p w14:paraId="11111111">${alt(outerParagraph, outerParagraph)}</w:p>`;
    const parser = new DocxParser(pkg(body));
    expect(parser.getDuplicateParaIds()).toEqual([]);
  });

  it('a repeat inside one branch is still a duplicate', () => {
    const body = `<w:p w14:paraId="11111111">${alt(textBox('22222222', 'a') + textBox('22222222', 'b'), textBox('22222222', 'a'))}</w:p>`;
    const parser = new DocxParser(pkg(body));
    const dups = parser.getDuplicateParaIds();
    expect(dups).toHaveLength(1);
    expect(dups[0].paraId).toBe('22222222');
    expect(dups[0].count).toBe(3);
  });

  it('a text box id repeated by a body paragraph is still a duplicate', () => {
    const body = `<w:p w14:paraId="11111111">${alt(textBox('22222222', 'a'), textBox('22222222', 'a'))}</w:p>` + textBox('22222222', 'c');
    const parser = new DocxParser(pkg(body));
    expect(parser.getDuplicateParaIds()).toHaveLength(1);
    expect(() => parser.rezip()).toThrow(/Duplicate Guard/);
  });

  it('two body paragraphs with one id are still refused (control)', () => {
    const parser = new DocxParser(pkg(textBox('55555555', 'a') + textBox('55555555', 'b')));
    expect(parser.getDuplicateParaIds()).toHaveLength(1);
  });
});
