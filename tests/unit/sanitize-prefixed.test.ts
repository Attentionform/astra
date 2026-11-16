/**
 * DN3 sanitizer ingress and post-conditions. A package that binds WordprocessingML to another prefix gets the
 * same scrub as the usual form, across body, header, footnotes and comments; a scrub that cannot do what an option
 * says fails instead of reporting success.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { createBlankDocxBuffer, sanitizeDocxBuffer } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rename = (xml: string, to = 'x'): string =>
  xml.replace(/(<\/?)w:/g, `$1${to}:`).replace(/(\s)w:([A-Za-z])/g, `$1${to}:$2`).replace(/xmlns:w=/g, `xmlns:${to}=`);
const ROOT = (tag: string, body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${tag} xmlns:w="${W}">${body}</w:${tag}>`;

function build(prefix: string | null): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const swap = (xml: string): string => (prefix ? rename(xml, prefix) : xml);
  const doc = zip.readAsText('word/document.xml').replace('<w:body>',
    '<w:body><w:p w:rsidR="00AA11BB"><w:ins w:id="9" w:author="Alice Secret" w:date="2024-01-01T00:00:00Z"><w:r><w:t>body bob@example.com</w:t></w:r></w:ins>'
    + '<w:commentRangeStart w:id="0"/><w:r><w:t>anchored</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>');
  zip.updateFile('word/document.xml', Buffer.from(swap(doc)));
  zip.addFile('word/header1.xml', Buffer.from(swap(ROOT('hdr', '<w:p w:rsidR="00CC22DD"><w:ins w:id="20" w:author="Alice Secret" w:date="2024-01-01T00:00:00Z"><w:r><w:t>header</w:t></w:r></w:ins></w:p>'))));
  zip.addFile('word/footnotes.xml', Buffer.from(swap(ROOT('footnotes', '<w:footnote w:id="1"><w:p w:rsidR="00EE33FF"><w:del w:id="21" w:author="Alice Secret" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>note</w:delText></w:r></w:del></w:p></w:footnote>'))));
  zip.addFile('word/comments.xml', Buffer.from(swap(ROOT('comments', '<w:comment w:id="0" w:author="Alice Secret" w:initials="AS" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>secret comment</w:t></w:r></w:p></w:comment>'))));
  return zip.toBuffer();
}
const parts = (buf: Buffer): string => new AdmZip(buf).getEntries().filter((e) => /\.xml$/.test(e.entryName)).map((e) => `${e.entryName}\n${e.getData().toString('utf8')}`).join('\n');

describe('sanitizer on a package with another WordprocessingML prefix', () => {
  const opts = { purgeRSIDs: true, purgeAuthors: true, purgePII: true, acceptAllRevisions: true, stripComments: true };
  const canonical = sanitizeDocxBuffer(build(null), opts);
  for (const prefix of ['x', 'q9', 'wml']) {
    it(`prefix "${prefix}": same report as the usual form, nothing left behind`, () => {
      const r = sanitizeDocxBuffer(build(prefix), opts);
      const strip = (rep: typeof r.report) => ({ ...rep, filesTouched: [...rep.filesTouched].sort() });
      expect(strip(r.report)).toEqual(strip(canonical.report));
      const all = parts(r.buffer);
      expect(all).not.toContain('Alice Secret');
      expect(all).not.toContain('bob@example.com');
      expect(all).not.toContain('secret comment');
      expect(all).not.toMatch(/rsid[A-Z]/);
      expect(all).not.toMatch(/<\w+:(ins|del) /);
      expect(all).not.toMatch(/commentRange|commentReference/);
    });
  }

  it('the old behaviour was silent: this is the defect the ingress closes (rsid and author counts equal the usual form)', () => {
    const r = sanitizeDocxBuffer(build('x'), opts);
    expect(r.report.removedRSIDs).toBe(canonical.report.removedRSIDs);
    expect(r.report.removedRSIDs).toBeGreaterThanOrEqual(3);
  });

  it('a conflicting package is refused and nothing is returned', () => {
    const zip = new AdmZip(build('x'));
    const doc = zip.readAsText('word/document.xml').replace('<x:body>', `<x:body><e xmlns:w="urn:other"><x:p/></e>`);
    zip.updateFile('word/document.xml', Buffer.from(doc));
    expect(() => sanitizeDocxBuffer(zip.toBuffer(), opts)).toThrow(/cannot be canonicalised safely/);
  });
});

describe('sanitizer post-conditions', () => {
  it('fails, naming option and part, when an author name survives in a part the scrub does not reach', () => {
    const zip = new AdmZip(build(null));
    zip.addFile('word/glossary/odd.xml', Buffer.from(ROOT('x', '<w:p><w:ins w:id="1" w:author="Survivor" w:date="2024-01-01T00:00:00Z"/></w:p>')));
    expect(() => sanitizeDocxBuffer(zip.toBuffer(), { purgeAuthors: true })).toThrow(/purgeAuthors.*word\/glossary\/odd\.xml/);
  });
});
