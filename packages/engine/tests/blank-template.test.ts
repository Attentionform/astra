import { afterEach, describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, composeDocument, createBlankDocxBuffer, validateDocxBuffer } from '../src/internal';

const part = (buf: Buffer, name: string) => new AdmZip(buf).readAsText(name);
const texts = (p: DocxParser) => p.getStructuralSummary().nodes.map((n) => n.text ?? '');

describe('the blank template is blank', () => {
  it('has one empty paragraph and no visible text', () => {
    const parser = DocxParser.createNewDocument('blank');
    expect(parser.getStructuralSummary().nodes).toHaveLength(1);
    expect(texts(parser)).toEqual(['']);
    expect(part(parser.rezip(), 'word/document.xml')).not.toContain('<w:t');
  });

  it('createBlankDocxBuffer gives the same document (one path, not two)', () => {
    const body = part(createBlankDocxBuffer('blank'), 'word/document.xml');
    expect(body).not.toContain('<w:t');
    expect(body).not.toContain('New Document');
    expect(body).toBe(part(DocxParser.createNewDocument('blank').rezip(), 'word/document.xml'));
  });

  it('carries no title and no author (never "ASTra")', () => {
    const core = part(DocxParser.createNewDocument('blank').rezip(), 'docProps/core.xml');
    expect(core).toContain('<dc:title></dc:title>');
    expect(core).toContain('<dc:creator></dc:creator>');
    expect(core).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');
  });

  it('validates', async () => {
    const res = await validateDocxBuffer(DocxParser.createNewDocument('blank').rezip());
    expect(res.ok, JSON.stringify(res.errors)).toBe(true);
  });
});

describe('ASTRA_AUTHOR names the author of a new file', () => {
  const saved = process.env.ASTRA_AUTHOR;
  afterEach(() => {
    if (saved === undefined) delete process.env.ASTRA_AUTHOR;
    else process.env.ASTRA_AUTHOR = saved;
  });
  it('writes the configured author as creator', () => {
    process.env.ASTRA_AUTHOR = 'Ada Lovelace';
    expect(part(DocxParser.createNewDocument('blank').rezip(), 'docProps/core.xml')).toContain('<dc:creator>Ada Lovelace</dc:creator>');
  });
});

describe('compose on a blank document', () => {
  it('replaces the empty paragraph in the session that created it', async () => {
    const parser = DocxParser.createNewDocument('blank');
    expect((await composeDocument(parser, '# One\n\nBody.\n')).success).toBe(true);
    expect(texts(parser)).toEqual(['One', 'Body.']);
  });

  it('replaces it after the file was saved and opened again (no leading empty paragraph)', async () => {
    const reopened = new DocxParser(Buffer.from(DocxParser.createNewDocument('blank').rezip()));
    expect((await composeDocument(reopened, '# One\n\nBody.\n')).success).toBe(true);
    expect(texts(reopened)).toEqual(['One', 'Body.']);
  });

  it('keeps a lone empty paragraph that holds a section break, and appends', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
    const xml = zip.readAsText('word/document.xml').replace(
      /<w:p\b[^>]*>\s*<w:pPr><w:pStyle w:val="Normal"\/><\/w:pPr>\s*<\/w:p>/,
      '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>'
    );
    expect(xml).toContain('<w:pStyle w:val="Normal"/><w:sectPr>');
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const reopened = new DocxParser(zip.toBuffer());
    expect((await composeDocument(reopened, 'Appended.\n')).success).toBe(true);
    expect(part(reopened.rezip(), 'word/document.xml')).toContain('<w:pStyle w:val="Normal"/><w:sectPr>');
    expect(texts(reopened)).toContain('Appended.');
  });

  it('keeps a lone empty paragraph that holds a bookmark, and appends', async () => {
    const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
    const xml = zip.readAsText('word/document.xml').replace(
      /<w:p\b[^>]*>\s*<w:pPr><w:pStyle w:val="Normal"\/><\/w:pPr>\s*<\/w:p>/,
      '<w:p><w:bookmarkStart w:id="7" w:name="keep_me"/><w:bookmarkEnd w:id="7"/></w:p>'
    );
    expect(xml).toContain('keep_me');
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const reopened = new DocxParser(zip.toBuffer());
    expect((await composeDocument(reopened, 'Appended.\n')).success).toBe(true);
    expect(part(reopened.rezip(), 'word/document.xml')).toContain('keep_me');
  });
});
