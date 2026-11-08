import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { sanitizeDocxBuffer } from '../../packages/engine/src/sanitizer';

const FIX = (name: string) => readFileSync(join(__dirname, '..', 'fixtures', 'word-revisions', `${name}.docx`));
const names = (buf: Buffer) => new AdmZip(buf).getEntries().map(e => e.entryName);
const allText = (buf: Buffer) => new AdmZip(buf).getEntries().map(e => e.getData().toString('latin1')).join('\n');
const run = (buf: Buffer, o: Record<string, unknown>) => sanitizeDocxBuffer(buf, o as never).buffer;

const S = 'w:author="Alice" w:date="2026-01-01T00:00:00Z"';
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const comment = (id: number, text: string, para: string) => `<w:comment w:id="${id}" w:author="Alice" w:date="2026-01-01T00:00:00Z" w:initials="A"><w:p w14:paraId="${para}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;

function build(body: string, comments: { id: number; text: string; para: string }[]): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  let rels = zip.readAsText('word/_rels/document.xml.rels');
  let types = zip.readAsText('[Content_Types].xml');
  const add = (part: string, xml: string, rel: string, ct: string) => {
    zip.addFile(`word/${part}`, Buffer.from(xml));
    rels = rels.replace('</Relationships>', `<Relationship Id="rIdX${part.replace(/\W/g, '')}" Type="${rel}" Target="${part}"/></Relationships>`);
    types = types.replace('</Types>', `<Override PartName="/word/${part}" ContentType="${ct}"/></Types>`);
  };
  add('comments.xml', `<w:comments ${NS}>${comments.map(c => comment(c.id, c.text, c.para)).join('')}</w:comments>`, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  add('commentsExtended.xml', `<w15:commentsEx ${NS}>${comments.map(c => `<w15:commentEx w15:paraId="${c.para}" w15:done="0"/>`).join('')}</w15:commentsEx>`, 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended', 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
  zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
  zip.updateFile('[Content_Types].xml', Buffer.from(types));
  return zip.toBuffer();
}
const anchored = (id: number, inner: string) => `<w:commentRangeStart w:id="${id}"/>${inner}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
const deletedParagraph = (id: number, text: string, cid: number) =>
  `<w:p><w:pPr><w:rPr><w:del w:id="${id}" ${S}/></w:rPr></w:pPr><w:commentRangeStart w:id="${cid}"/><w:del w:id="${id + 1}" ${S}><w:r><w:delText>${text}</w:delText></w:r></w:del><w:commentRangeEnd w:id="${cid}"/><w:del w:id="${id + 2}" ${S}><w:r><w:commentReference w:id="${cid}"/></w:r></w:del></w:p>`;
const live = (text: string, cid?: number) => `<w:p>${cid === undefined ? `<w:r><w:t>${text}</w:t></w:r>` : anchored(cid, `<w:r><w:t>${text}</w:t></w:r>`)}</w:p>`;

describe('stripComments removes every comment part, not only comments.xml', () => {
  it('leaves no comment part, relationship, content type or comment timestamp in a Word 2016 file', () => {
    const out = run(FIX('delete-paragraph-comment-thread'), { acceptAllRevisions: true, stripComments: true });
    expect(names(out).filter(n => /comments/i.test(n))).toEqual([]);
    const zip = new AdmZip(out);
    expect(zip.readAsText('word/_rels/document.xml.rels')).not.toMatch(/comments/i);
    expect(zip.readAsText('[Content_Types].xml')).not.toMatch(/comments/i);
    expect(allText(out)).not.toContain('dateUtc');
  });

  it('the stripped file is valid', async () => {
    const out = run(FIX('delete-paragraph-comment-thread'), { acceptAllRevisions: true, stripComments: true });
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });
});

describe('accepting revisions deletes the comments of a removed paragraph (FX-sanitizer-comments)', () => {
  for (const name of ['delete-paragraph-comment-thread', 'delete-heading-bookmark-comment']) {
    it(`${name}: no comment text, comment part, timestamp or id survives, and the file is valid`, async () => {
      const original = FIX(name);
      const texts = (allText(original).match(/<w:t[^>]*>[^<]*<\/w:t>/g) ?? []);
      expect(texts.length).toBeGreaterThan(0);
      const commentWords = new AdmZip(original).readAsText('word/comments.xml').match(/<w:t[^>]*>([^<]+)<\/w:t>/g)!.map(t => t.replace(/<[^>]+>/g, ''));
      const out = run(original, { acceptAllRevisions: true });
      for (const word of commentWords) expect(allText(out), word).not.toContain(word);
      expect(names(out).filter(n => /comments/i.test(n))).toEqual([]);
      expect(allText(out)).not.toContain('dateUtc');
      expect((await validateDocxBuffer(out)).ok).toBe(true);
    });
  }

  it('a comment on a paragraph that stays is kept, with its extended data', async () => {
    const buf = build(deletedParagraph(10, 'gone text', 0) + live('kept text', 1) + live('tail'), [
      { id: 0, text: 'comment on deleted text', para: '11111111' },
      { id: 1, text: 'comment on kept text', para: '22222222' },
    ]);
    const out = run(buf, { acceptAllRevisions: true });
    const zip = new AdmZip(out);
    const comments = zip.readAsText('word/comments.xml');
    expect(comments).toContain('comment on kept text');
    expect(comments).not.toContain('comment on deleted text');
    const ext = zip.readAsText('word/commentsExtended.xml');
    expect(ext).toContain('22222222');
    expect(ext).not.toContain('11111111');
    expect(zip.readAsText('word/document.xml')).toContain('w:id="1"');
    expect((await validateDocxBuffer(out)).ok).toBe(true);
  });

  it('with stripComments every comment goes, as before', () => {
    const buf = build(deletedParagraph(10, 'gone text', 0) + live('kept text', 1), [
      { id: 0, text: 'a', para: '11111111' }, { id: 1, text: 'b', para: '22222222' },
    ]);
    expect(names(run(buf, { acceptAllRevisions: true, stripComments: true })).filter(n => /comments/i.test(n))).toEqual([]);
  });

  it('a comment that is also anchored outside the removed paragraph makes the sanitizer refuse, and names the way out', () => {
    // range starts in the deleted paragraph and ends in the next, live one
    const body = `<w:p><w:pPr><w:rPr><w:del w:id="10" ${S}/></w:rPr></w:pPr><w:commentRangeStart w:id="0"/><w:del w:id="11" ${S}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`
      + `<w:p><w:r><w:t>live</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
    const buf = build(body, [{ id: 0, text: 'spanning', para: '11111111' }]);
    expect(() => run(buf, { acceptAllRevisions: true })).toThrow(/E_VALIDATION[\s\S]*revision markup remains[\s\S]*stripComments/);
  });

  it('a document with comments and no revisions is unchanged by acceptAllRevisions', () => {
    const buf = build(live('kept text', 0), [{ id: 0, text: 'keep me', para: '11111111' }]);
    const out = run(buf, { acceptAllRevisions: true });
    expect(new AdmZip(out).readAsText('word/comments.xml')).toContain('keep me');
  });
});
