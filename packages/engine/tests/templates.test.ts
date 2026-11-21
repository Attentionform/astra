import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, TEMPLATE_NAMES, composeDocument, validateDocxBuffer } from '../src/internal';
import { findRegion, findSlot, hashNodes } from '../src/compose/template-canonical';
import { TEMPLATE_MANIFEST } from '../src/compose/template-manifest.generated';
import { DOMParser } from '../src/ooxml/dom-parser';
import { contrastChecks } from '../scripts/templates/design';
import { DESIGN } from '../scripts/templates/themes';

const part = (buf: Buffer, name: string): string => new AdmZip(buf).readAsText(name);
const SIZES = ['letter', 'a4'] as const;
const DESIGNED = TEMPLATE_NAMES.filter((n) => n !== 'blank');

/** Visible text of every paragraph, field results included. */
const textOf = (xml: string): string => Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g), (m) => m[1]).join('\n');

describe('every template, both paper sizes', () => {
  for (const name of TEMPLATE_NAMES) {
    for (const size of SIZES) {
      it(`${name} (${size}) is a valid, unsigned starting document`, async () => {
        const buf = DocxParser.createNewDocument(name, { pageSize: size }).rezip();
        const res = await validateDocxBuffer(buf);
        expect(res.ok, JSON.stringify(res.errors)).toBe(true);
        const core = part(buf, 'docProps/core.xml');
        expect(core).toContain('<dc:creator></dc:creator>');
        expect(core).toContain('<dc:title></dc:title>');
        expect(core).not.toMatch(/ASTra/);
        const doc = part(buf, 'word/document.xml');
        expect(doc).toMatch(size === 'a4' ? /<w:pgSz w:w="11906" w:h="16838"\/>/ : /<w:pgSz w:w="12240" w:h="15840"\/>/);
        expect(part(buf, 'docProps/custom.xml')).toContain(`${name}@1`);
      });
    }
  }
});

describe('starters hold placeholders, not invented facts', () => {
  for (const name of TEMPLATE_NAMES) {
    it(`${name}: no figures, years, currency, percentages or product names`, () => {
      const buf = DocxParser.createNewDocument(name).rezip();
      for (const file of ['word/document.xml', 'word/header1.xml', 'word/footer1.xml']) {
        let xml: string;
        try {
          xml = part(buf, file);
        } catch {
          continue;
        }
        // Field instructions and cached field results are not content.
        const visible = textOf(xml.replace(/<w:instrText[^>]*>[^<]*<\/w:instrText>/g, '').replace(/<w:fldChar[^>]*\/>/g, ''));
        const withoutFieldResults = visible.replace(/^1$/gm, '').replace(/Table \d+\./g, 'Table N.');
        expect(withoutFieldResults, file).not.toMatch(/\d{4}/);
        expect(withoutFieldResults, file).not.toMatch(/[$€£%]/);
        expect(withoutFieldResults, file).not.toMatch(/ASTra/);
      }
    });
  }

  it('blank is one empty paragraph and nothing else', () => {
    const buf = DocxParser.createNewDocument('blank').rezip();
    expect(textOf(part(buf, 'word/document.xml'))).toBe('');
    expect(part(buf, 'word/document.xml').match(/<w:p[ >]/g)).toHaveLength(1);
  });
});

describe('every style a template body uses is defined in that template', () => {
  for (const name of TEMPLATE_NAMES) {
    it(name, () => {
      const buf = DocxParser.createNewDocument(name).rezip();
      const styles = part(buf, 'word/styles.xml');
      const defined = new Set(Array.from(styles.matchAll(/w:styleId="([^"]+)"/g), (m) => m[1]));
      for (const file of ['word/document.xml', 'word/header1.xml', 'word/footer1.xml']) {
        let xml: string;
        try {
          xml = part(buf, file);
        } catch {
          continue;
        }
        for (const m of xml.matchAll(/<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"/g)) {
          expect(defined.has(m[1]), `${name}: ${m[1]} used in ${file}`).toBe(true);
        }
      }
    });
  }
  it('the headings, lists, quote, code, caption and table styles exist', () => {
    const styles = part(DocxParser.createNewDocument('blank').rezip(), 'word/styles.xml');
    for (const id of ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'ListBullet', 'ListNumber', 'Quote', 'Code', 'Caption', 'Title', 'Subtitle', 'AstraTable', 'Placeholder', 'Callout', 'TOC1', 'Header', 'Footer', 'TableGrid', 'FootnoteText', 'FootnoteReference', 'EndnoteText', 'EndnoteReference', 'Hyperlink']) {
      expect(styles, id).toContain(`w:styleId="${id}"`);
    }
  });
});

describe('the design palette meets WCAG AA', () => {
  for (const c of contrastChecks(DESIGN)) {
    it(`${c.pair}: ${c.ratio}:1`, () => expect(c.ok).toBe(true));
  }
});

describe('the shipped manifest matches the shipped files', () => {
  for (const name of DESIGNED) {
    for (const size of SIZES) {
      it(`${name} (${size}): body hash, in the session and after save and reopen`, () => {
        const entry = TEMPLATE_MANIFEST[`${name}@1`];
        expect(entry, 'manifest entry').toBeDefined();
        const fresh = DocxParser.createNewDocument(name, { pageSize: size });
        for (const parser of [fresh, new DocxParser(Buffer.from(fresh.rezip()))]) {
          const body = parser.getDocumentDom().getElementsByTagName('w:body')[0];
          const region = findRegion(body);
          expect(region, 'region markers').not.toBeNull();
          expect(entry.region).toContain(hashNodes(region!.nodes));
        }
      });
    }
  }
  it('the canonical hash is the same when the XML is only re-indented', () => {
    const buf = DocxParser.createNewDocument('memo').rezip();
    const xml = part(buf, 'word/document.xml');
    const spaced = xml.replace(/></g, '>\n  <');
    const a = findRegion(new DOMParser().parseFromString(xml, 'text/xml').getElementsByTagName('w:body')[0]);
    const b = findRegion(new DOMParser().parseFromString(spaced, 'text/xml').getElementsByTagName('w:body')[0]);
    expect(hashNodes(a!.nodes)).toBe(hashNodes(b!.nodes));
  });
});

describe('unknown names and sizes', () => {
  it('an unknown template is an error that lists the valid names', () => {
    expect(() => DocxParser.createNewDocument('corporate')).toThrow(/unknown template 'corporate'.*blank, report, memo/);
    expect(() => DocxParser.createNewDocument('api-spec')).toThrow(/unknown template/);
  });
  it('an unknown page size is an error', () => {
    expect(() => DocxParser.createNewDocument('report', { pageSize: 'legal' as never })).toThrow(/page size/);
  });
});

const MD = '# Quarterly Review\n\nFirst paragraph of the body.\n\n## Detail\n\nSecond paragraph.\n';

describe('compose onto an untouched template replaces the placeholder body and keeps the furniture', () => {
  for (const name of DESIGNED) {
    it(name, async () => {
      const parser = DocxParser.createNewDocument(name);
      const before = parser.rezip();
      const res = await composeDocument(parser, MD);
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
      const after = parser.rezip();
      const doc = part(after, 'word/document.xml');
      const text = textOf(doc);
      expect(text).toContain('First paragraph of the body.');
      expect(text).toContain('Second paragraph.');
      // The placeholder body is gone: no bracketed placeholder remains inside the composed region (furniture may keep its own).
      expect(doc).not.toContain('astra_body_start');
      expect(doc).not.toContain('astra_body_end');
      expect(part(after, 'docProps/custom.xml')).toContain(`${name}@1:composed`);
      // The title slot took the heading; the heading is not repeated in the body.
      expect(text.match(/Quarterly Review/g)).toHaveLength(1);
      // Furniture parts are byte-identical.
      for (const f of ['word/styles.xml', 'word/numbering.xml', 'word/settings.xml', 'word/theme/theme1.xml', 'word/header1.xml', 'word/footer1.xml']) {
        const was = (() => {
          try {
            return part(before, f);
          } catch {
            return null;
          }
        })();
        if (was !== null) expect(part(after, f), f).toBe(was);
      }
      const res2 = await validateDocxBuffer(after);
      expect(res2.ok, JSON.stringify(res2.errors)).toBe(true);
    });
  }

  it('also after the file was saved and opened again', async () => {
    const reopened = new DocxParser(Buffer.from(DocxParser.createNewDocument('report').rezip()));
    const res = await composeDocument(reopened, MD);
    expect(res.success).toBe(true);
    expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
    expect(textOf(part(reopened.rezip(), 'word/document.xml'))).not.toContain('Summarise the purpose');
  });

  it('a second compose appends (the template is marked composed)', async () => {
    const parser = DocxParser.createNewDocument('report');
    await composeDocument(parser, MD);
    const res = await composeDocument(parser, 'Added later.\n');
    expect(res.success).toBe(true);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_ALREADY_COMPOSED'))).toBe(true);
    const text = textOf(part(parser.rezip(), 'word/document.xml'));
    expect(text).toContain('First paragraph of the body.');
    expect(text).toContain('Added later.');
  });

  it('a heading with inline formatting stays in the body (the slot only takes plain text)', async () => {
    const parser = DocxParser.createNewDocument('report');
    await composeDocument(parser, '# A **bold** title\n\nBody.\n');
    const doc = part(parser.rezip(), 'word/document.xml');
    expect(doc).toContain('[Report title]');
    expect(textOf(doc)).toContain('title');
  });
});

describe('compose never deletes what the user changed (R3: append plus a warning)', () => {
  const composeOn = async (parser: DocxParser) => composeDocument(parser, 'Appended text.\n');

  it('edited text inside the body: kept, TEMPLATE_REGION_EDITED', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('report').rezip());
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('[Describe the context, scope and sources the report draws on.]', 'My own background text.')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.success).toBe(true);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_REGION_EDITED'))).toBe(true);
    const text = textOf(part(parser.rezip(), 'word/document.xml'));
    expect(text).toContain('My own background text.');
    expect(text).toContain('Appended text.');
    expect(text).toContain('Summarise the purpose');
  });

  it('a deleted boundary bookmark: kept, TEMPLATE_BOOKMARKS_MISSING', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('memo').rezip());
    const doc = zip.readAsText('word/document.xml').replace(/<w:bookmarkStart w:id="3" w:name="astra_body_end"\/><w:bookmarkEnd w:id="3"\/>/, '');
    expect(doc).not.toContain('astra_body_end');
    zip.updateFile('word/document.xml', Buffer.from(doc));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_BOOKMARKS_MISSING'))).toBe(true);
    expect(textOf(part(parser.rezip(), 'word/document.xml'))).toContain('State in one or two sentences');
  });

  it('a removed marker property: an ordinary document, nothing deleted, no warning', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('memo').rezip());
    zip.updateFile('docProps/custom.xml', Buffer.from(part(zip.toBuffer(), 'docProps/custom.xml').replace(/<property .*<\/property>/, '')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.success).toBe(true);
    expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
    expect(textOf(part(parser.rezip(), 'word/document.xml'))).toContain('State in one or two sentences');
  });

  it('a corrupt marker property: TEMPLATE_MARKER_INVALID, nothing deleted', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('memo').rezip());
    zip.updateFile('docProps/custom.xml', Buffer.from(part(zip.toBuffer(), 'docProps/custom.xml').replace('memo@1', 'garbage value')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_MARKER_INVALID'))).toBe(true);
    expect(textOf(part(parser.rezip(), 'word/document.xml'))).toContain('State in one or two sentences');
  });

  it('an unknown design version: TEMPLATE_MANIFEST_MISSING, nothing deleted', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('memo').rezip());
    zip.updateFile('docProps/custom.xml', Buffer.from(part(zip.toBuffer(), 'docProps/custom.xml').replace('memo@1', 'memo@99')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_MANIFEST_MISSING'))).toBe(true);
    expect(textOf(part(parser.rezip(), 'word/document.xml'))).toContain('State in one or two sentences');
  });

  it('a self-consistent tampered body cannot claim to be untouched: the shipped manifest decides', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('memo').rezip());
    // The document is edited and keeps the marker and the bookmarks; only the manifest can tell.
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('[State in one or two sentences why you are writing and what you need.]', '[Precious user text]')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_REGION_EDITED'))).toBe(true);
    expect(textOf(part(parser.rezip(), 'word/document.xml'))).toContain('Precious user text');
  });

  it('edits to the furniture (cover text, header, styles, properties) do not stop the replacement', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('report').rezip());
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('[Client or audience]', 'Acme Ltd')));
    zip.updateFile('word/styles.xml', Buffer.from(zip.readAsText('word/styles.xml').replace('w:val="1F3A5F"', 'w:val="AA0000"')));
    zip.updateFile('docProps/core.xml', Buffer.from(zip.readAsText('docProps/core.xml').replace('<dc:subject></dc:subject>', '<dc:subject>Edited</dc:subject>')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeOn(parser);
    expect(res.success).toBe(true);
    expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
    const text = textOf(part(parser.rezip(), 'word/document.xml'));
    expect(text).toContain('Acme Ltd');
    expect(text).toContain('Appended text.');
    expect(text).not.toContain('Summarise the purpose');
  });

  it('a template option that names another template is an error and writes nothing', async () => {
    const parser = DocxParser.createNewDocument('memo');
    const before = parser.rezip();
    await expect(composeDocument(parser, 'x\n', { template: 'report' })).rejects.toThrow(/not a fresh report template/);
    expect(part(parser.rezip(), 'word/document.xml')).toBe(part(before, 'word/document.xml'));
  });
});

describe('a bookmark that someone added to the placeholder body is never deleted', () => {
  // Word adds _Toc bookmarks to headings and _Ref bookmarks to cross-reference targets by itself. A bookmark with one of those
  // names that did not come from Word must make the body count as edited, so that compose appends and deletes nothing.
  const withBookmarks = (xml: string, inject: (xml: string) => string): Buffer => {
    const zip = new AdmZip(DocxParser.createNewDocument('report').rezip());
    zip.updateFile('word/document.xml', Buffer.from(inject(xml)));
    return zip.toBuffer();
  };
  const firstHeadingRun = (xml: string): { before: string; after: string; start: number } => {
    const at = xml.indexOf('<w:r>', xml.indexOf('astra_body_start'));
    return { before: xml.slice(0, at), after: xml.slice(at), start: at };
  };
  const cases: Array<[string, (xml: string) => string]> = [
    ['an unreferenced _Toc pair inside one paragraph', (x) => { const f = firstHeadingRun(x); return `${f.before}<w:bookmarkStart w:id="901" w:name="_Toc987654"/>${f.after.replace('</w:p>', '<w:bookmarkEnd w:id="901"/></w:p>')}`; }],
    ['an unreferenced _Ref pair inside one paragraph', (x) => { const f = firstHeadingRun(x); return `${f.before}<w:bookmarkStart w:id="902" w:name="_Ref123456789"/>${f.after.replace('</w:p>', '<w:bookmarkEnd w:id="902"/></w:p>')}`; }],
    ['a _Toc pair that starts in one paragraph and ends in the next', (x) => { const f = firstHeadingRun(x); const i = f.after.indexOf('</w:p>'); const j = f.after.indexOf('</w:p>', i + 6); return `${f.before}<w:bookmarkStart w:id="903" w:name="_Toc555"/>${f.after.slice(0, j)}<w:bookmarkEnd w:id="903"/>${f.after.slice(j)}`; }],
    ['a _GoBack pair that wraps text', (x) => { const f = firstHeadingRun(x); return `${f.before}<w:bookmarkStart w:id="904" w:name="_GoBack"/>${f.after.replace('</w:p>', '<w:bookmarkEnd w:id="904"/></w:p>')}`; }],
  ];
  for (const [label, inject] of cases) {
    it(`${label}: compose appends and keeps it`, async () => {
      const buf = withBookmarks(part(DocxParser.createNewDocument('report').rezip(), 'word/document.xml'), inject);
      expect(part(buf, 'word/document.xml')).toMatch(/w:name="_(Toc|Ref|GoBack)/);
      const parser = new DocxParser(buf);
      const res = await composeDocument(parser, '# Title\n\nBody.\n');
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      expect(res.warnings.some((w) => w.startsWith('TEMPLATE_REGION_EDITED'))).toBe(true);
      expect(part(parser.rezip(), 'word/document.xml')).toMatch(/w:name="_(Toc|Ref|GoBack)/);
    });
  }
  // A link or a field outside the placeholder body points at the bookmark. Deleting the bookmark with the body would break it.
  const outside = (target: string, name: string, link: (name: string) => string): ((x: string) => string) => (x) => {
    const f = firstHeadingRun(x);
    const withBookmark = `${f.before}<w:bookmarkStart w:id="910" w:name="${name}"/>${f.after.replace('</w:p>', '<w:bookmarkEnd w:id="910"/></w:p>')}`;
    // The paragraph that holds the link goes before the first body paragraph, outside the replaceable region.
    const body = withBookmark.indexOf('<w:body>') + '<w:body>'.length;
    // A table of contents entry points at the bookmark as well, as it does for a heading that Word wrote.
    const tocEntry = `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="${name}" w:history="1"><w:r><w:t>Entry</w:t></w:r></w:hyperlink></w:p>`;
    return `${withBookmark.slice(0, body)}${tocEntry}<w:p>${link(target)}</w:p>${withBookmark.slice(body)}`;
  };
  const referenced: Array<[string, (x: string) => string]> = [
    ['a _Toc bookmark that a hyperlink in the text points at', outside('_Toc987654', '_Toc987654', (n) => `<w:hyperlink w:anchor="${n}" w:history="1"><w:r><w:t>see below</w:t></w:r></w:hyperlink>`)],
    ['a _Toc bookmark that a REF field in the text points at', outside('_Toc987654', '_Toc987654', (n) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ${n} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`)],
    ['a _Ref bookmark that a REF field in the text points at', outside('_Ref123456789', '_Ref123456789', (n) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF ${n} \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`)],
    ['a _GoBack bookmark that a hyperlink points at', outside('_GoBack', '_GoBack', (n) => `<w:hyperlink w:anchor="${n}" w:history="1"><w:r><w:t>back</w:t></w:r></w:hyperlink>`)],
    ['a _Toc bookmark that a simple REF field in the text points at', outside('_Toc987654', '_Toc987654', (n) => `<w:fldSimple w:instr=" REF ${n} \\h "><w:r><w:t>x</w:t></w:r></w:fldSimple>`)],
    ['a _Toc bookmark that a simple PAGEREF field in the text points at, beside a table of contents link', outside('_Toc987654', '_Toc987654', (n) => `<w:fldSimple w:instr=" PAGEREF ${n} \\h "><w:r><w:t>2</w:t></w:r></w:fldSimple>`)],
  ];
  for (const [label, inject] of referenced) {
    it(`${label}: compose appends and keeps both the bookmark and the link`, async () => {
      const buf = withBookmarks(part(DocxParser.createNewDocument('report').rezip(), 'word/document.xml'), inject);
      const parser = new DocxParser(buf);
      const res = await composeDocument(parser, '# Title\n\nBody.\n');
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      expect(res.warnings.some((w) => w.startsWith('TEMPLATE_REGION_EDITED'))).toBe(true);
      const out = part(parser.rezip(), 'word/document.xml');
      expect(out).toMatch(/<w:bookmarkStart [^>]*w:name="_(Toc987654|Ref123456789|GoBack)"/);
      expect(out).toMatch(/w:anchor="_(Toc987654|GoBack)"|REF _(Toc987654|Ref123456789)|w:instr=" (PAGE)?REF _Toc987654/);
    });
  }
  it('a _Toc bookmark that a field in a header points at: compose appends and keeps it', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('report').rezip());
    const xml = part(zip.toBuffer(), 'word/document.xml');
    const f = firstHeadingRun(xml);
    const withBookmark = `${f.before}<w:bookmarkStart w:id="920" w:name="_Toc987654"/>${f.after.replace('</w:p>', '<w:bookmarkEnd w:id="920"/></w:p>')}`;
    const body = withBookmark.indexOf('<w:body>') + '<w:body>'.length;
    const tocEntry = '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc987654" w:history="1"><w:r><w:t>Entry</w:t></w:r></w:hyperlink></w:p>';
    zip.updateFile('word/document.xml', Buffer.from(`${withBookmark.slice(0, body)}${tocEntry}${withBookmark.slice(body)}`));
    const header = part(zip.toBuffer(), 'word/header1.xml');
    zip.updateFile('word/header1.xml', Buffer.from(header.replace(/<\/w:p>/, '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Toc987654 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeDocument(parser, '# Title\n\nBody.\n');
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_REGION_EDITED'))).toBe(true);
    expect(part(parser.rezip(), 'word/document.xml')).toMatch(/w:name="_Toc987654"/);
  });
  it('an empty _GoBack pair, which Word adds where the cursor was, is still ignored', async () => {
    const buf = withBookmarks(part(DocxParser.createNewDocument('report').rezip(), 'word/document.xml'), (x) => { const f = firstHeadingRun(x); return `${f.before}<w:bookmarkStart w:id="905" w:name="_GoBack"/><w:bookmarkEnd w:id="905"/>${f.after}`; });
    const parser = new DocxParser(buf);
    const res = await composeDocument(parser, '# Title\n\nBody.\n');
    expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
  });
});

describe('a leading heading reads Markdown entities as the compiler does', () => {
  const heading = async (template: string, markdown: string): Promise<string> => {
    const parser = new DocxParser(DocxParser.createNewDocument(template).rezip());
    const res = await composeDocument(parser, markdown);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const xml = part(parser.rezip(), 'word/document.xml');
    const match = /Heading1|Title/.test(xml) ? Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g), (m) => m[1]) : [];
    return match.join('|');
  };
  const entities: Array<[string, string]> = [
    ['# Sales &amp; Growth', 'Sales &amp; Growth'],
    ['# R &copy; 2026', 'R © 2026'],
    ['# Code &#65;&#x42;', 'Code AB'],
    ['# AT&T', 'AT&amp;T'],
    ['# Q&A &unknownname; end', 'Q&amp;A &amp;unknownname; end'],
  ];
  for (const [source, expectedXml] of entities) {
    it(`${source} gives the same title on a report as on a blank document`, async () => {
      const md = `${source}\n\nBody.\n`;
      const onBlank = await heading('blank', md);
      const onReport = await heading('report', md);
      expect(onBlank).toContain(expectedXml);
      expect(onReport).toContain(expectedXml);
      expect(onReport).not.toMatch(/&amp;amp;|&amp;copy;|&amp;#/);
    });
  }
});

describe('input that is only a heading, and an explicit anchor', () => {
  const visible = (buf: Buffer): string[] => Array.from(part(buf, 'word/document.xml').matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g), (m) => m[1]);
  for (const name of DESIGNED.filter((n) => TEMPLATE_MANIFEST[`${n}@1`]?.slots.title !== undefined)) {
    it(`${name}: a heading that is the whole input fills the title slot and removes the placeholder body`, async () => {
      const fresh = DocxParser.createNewDocument(name);
      const slot = findSlot(fresh.getDocumentDom().getElementsByTagName('w:body')[0], 'title');
      const placeholder = (slot?.nodes ?? []).map((n) => n.textContent ?? '').join('').trim();
      expect(placeholder, 'the template has a title placeholder').not.toBe('');
      const parser = new DocxParser(fresh.rezip());
      const res = await composeDocument(parser, '# Quarterly Review\n');
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
      const out = parser.rezip();
      const text = visible(out);
      expect(text.filter((x) => x.includes('Quarterly Review'))).toHaveLength(1);
      expect(text.join('')).not.toContain(placeholder);
      const valid = await validateDocxBuffer(out);
      expect(valid.ok, JSON.stringify(valid.errors)).toBe(true);
      const again = await composeDocument(parser, 'Next.\n');
      expect(again.warnings.some((w) => w.startsWith('TEMPLATE_ALREADY_COMPOSED'))).toBe(true);
    });
  }
  it('input with no content block at all is still refused', async () => {
    await expect(composeDocument(new DocxParser(DocxParser.createNewDocument('report').rezip()), '\n')).rejects.toThrow(/no content blocks found/);
  });
  it('an explicit anchor on a blank document appends and keeps the blank paragraph and its id', async () => {
    const parser = new DocxParser(DocxParser.createNewDocument('blank').rezip());
    const ids = Array.from(parser.nodeMap.keys());
    expect(ids).toHaveLength(1);
    const res = await composeDocument(parser, 'Hello.\n', { anchor: ids[0], position: 'after' });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(parser.nodeMap.has(ids[0])).toBe(true);
    expect(visible(parser.rezip())).toContain('Hello.');
    // The id is still usable by a later operation.
    const next = await composeDocument(parser, 'After the anchor.\n', { anchor: ids[0], position: 'after' });
    expect(next.success, JSON.stringify(next.errors)).toBe(true);
  });
});
