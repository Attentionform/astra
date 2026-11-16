/**
 * XML allows single-quoted attribute values and spaces around "=". The scrubs and the post-conditions used to
 * match double-quoted values only, so a file written that way kept its authors and revision-save ids while the
 * report said ok.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer, sanitizeDocxBuffer } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const ROOT = (tag: string, body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} xmlns:w="${W}" xmlns:w15="${W15}">${body}</w:${tag}>`;

/** `style`: how every attribute value is written. */
type Style = (name: string, value: string) => string;
const styles: Record<string, Style> = {
  double: (n, v) => `${n}="${v}"`,
  single: (n, v) => `${n}='${v}'`,
  'single, spaced': (n, v) => `${n} = '${v}'`,
  'double, spaced': (n, v) => `${n}\n=\n"${v}"`,
};

function build(q: Style): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const doc = zip.readAsText('word/document.xml').replace('<w:body>',
    `<w:body><w:p ${q('w:rsidR', 'A1B2C3D4')} ${q('w:rsidRDefault', 'A1B2C3D5')}><w:ins ${q('w:id', '42')} ${q('w:author', 'Alice Secret')} ${q('w:date', '2024-01-01T00:00:00Z')}><w:r><w:t>kept text</w:t></w:r></w:ins></w:p>`);
  zip.updateFile('word/document.xml', Buffer.from(doc));
  zip.addFile('word/comments.xml', Buffer.from(ROOT('comments', `<w:comment ${q('w:id', '0')} ${q('w:author', 'Bob Secret')} ${q('w:initials', 'BS')}><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment>`)));
  zip.addFile('word/people.xml', Buffer.from(ROOT('people', `<w15:person ${q('w15:author', 'Carol Secret')}><w15:presenceInfo ${q('w15:providerId', 'None')} ${q('w15:userId', 'Carol Secret')}/></w15:person>`).replace('<w:people', '<w15:people').replace('</w:people>', '</w15:people>')));
  const settings = zip.readAsText('word/settings.xml');
  zip.updateFile('word/settings.xml', Buffer.from(settings.replace(/<\/w:settings>\s*$/, `<w:rsids><w:rsidRoot ${q('w:val', '00A1B2C3')}/><w:rsid ${q('w:val', '00D4E5F6')}/></w:rsids></w:settings>`)));
  return zip.toBuffer();
}
const allXml = (buf: Buffer): string => new AdmZip(buf).getEntries().filter((e) => /\.xml$/.test(e.entryName)).map((e) => e.getData().toString('utf8')).join('\n');

describe('sanitize with every XML spelling of an attribute', () => {
  for (const [label, q] of Object.entries(styles)) {
    it(`${label}: authors and revision-save ids are removed, text is kept`, () => {
      const r = sanitizeDocxBuffer(build(q), { purgeRSIDs: true, purgeAuthors: true, purgePII: false, acceptAllRevisions: false, stripComments: false });
      expect(r.report.ok).toBe(true);
      const xml = allXml(r.buffer);
      for (const secret of ['Alice Secret', 'Bob Secret', 'Carol Secret', 'A1B2C3D4', 'A1B2C3D5', '00A1B2C3', '00D4E5F6']) {
        expect(xml, secret).not.toContain(secret);
      }
      expect(xml).toContain('kept text');
      expect(xml).not.toMatch(/w:initials\s*=\s*["']BS["']/);
    });
  }

  for (const [label, q] of Object.entries(styles)) {
    it(`${label}: an e-mail address in alternative text and in a person id is redacted by purgePII`, () => {
      const zip = new AdmZip(build(q));
      const doc = zip.readAsText('word/document.xml').replace('<w:body>',
        `<w:body><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:docPr ${q('id', '1')} ${q('name', 'Picture 1')} ${q('descr', 'photo by jane.doe@example.com')}/></wp:inline></w:drawing></w:r></w:p>`);
      zip.updateFile('word/document.xml', Buffer.from(doc));
      const r = sanitizeDocxBuffer(zip.toBuffer(), { purgePII: true, purgeAuthors: false, purgeRSIDs: false, acceptAllRevisions: false, stripComments: false });
      expect(r.report.ok).toBe(true);
      expect(allXml(r.buffer)).not.toContain('jane.doe@example.com');
    });
  }
});

describe('example markup inside an XML comment or CDATA section', () => {
  it('is not mistaken for a comment anchor, a revision or a leftover when stripping comments', () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const doc = zip.readAsText('word/document.xml').replace('<w:body>',
      '<w:body><!-- example: <w:commentReference w:id="1"/> and <w:ins w:id="5"> --><w:p><w:r><w:t>kept</w:t></w:r></w:p><w:p><w:r><w:t><![CDATA[<w:commentRangeStart w:id="2"/>]]></w:t></w:r></w:p>');
    zip.updateFile('word/document.xml', Buffer.from(doc));
    const r = sanitizeDocxBuffer(zip.toBuffer(), { stripComments: true, acceptAllRevisions: true, purgeAuthors: true, purgeRSIDs: true });
    expect(r.report.ok).toBe(true);
    expect(allXml(r.buffer)).toContain('kept');
  });
});

describe('example attributes inside an XML comment or CDATA section of a part the scrub does not rewrite', () => {
  for (const [label, part] of [['a custom XML part', 'word/customXml/item9.xml'], ['an unlisted word part', 'word/notes-example.xml']] as const) {
    it(`${label}: text such as w:author="Alice" or w:rsidR="1234" is not treated as an attribute`, () => {
      const zip = new AdmZip(createBlankDocxBuffer('blank'));
      zip.addFile(part, Buffer.from('<root><!-- example: <w:p w:rsidR="1234" w:author="Alice"/> --><note><![CDATA[<w:ins w:author="Alice" w:rsidR="5678"/>]]></note></root>'));
      const r = sanitizeDocxBuffer(zip.toBuffer(), { purgeRSIDs: true, purgeAuthors: true, purgePII: false, acceptAllRevisions: false, stripComments: false });
      expect(r.report.ok).toBe(true);
    });
  }
});
