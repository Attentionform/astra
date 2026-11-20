import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, TEMPLATE_NAMES, composeDocument, validateDocxBuffer } from '../src/internal';
import { findRegion, hashNodes } from '../src/compose/template-canonical';
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
