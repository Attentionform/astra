import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, applyDocumentIdentity, deriveTitleFromContent, isPlaceholderTitle, DERIVED_TITLE_MAX } from '../src/index';
import { composeDocument } from '../src/compose/markdown';

const coreXml = (parser: DocxParser) => new AdmZip(parser.rezip()).readAsText('docProps/core.xml');
const dcTitle = (parser: DocxParser) => /<dc:title>([^<]*)<\/dc:title>/.exec(coreXml(parser))?.[1];
const dcCreator = (parser: DocxParser) => /<dc:creator>([^<]*)<\/dc:creator>/.exec(coreXml(parser))?.[1];

describe('placeholder titles', () => {
  it('knows the template names and empty text', () => {
    for (const t of ['Blank Document', 'Report', 'Memo', 'Letter', 'Api Spec', 'Bar Chart 3d', '', '  ', undefined]) {
      expect(isPlaceholderTitle(t), String(t)).toBe(true);
    }
    expect(isPlaceholderTitle('MCP Engine Evaluation 2026')).toBe(false);
    expect(isPlaceholderTitle('My Report')).toBe(false);
  });
});

describe('the title is written from the content when it is still a placeholder', () => {
  it('compose: dc:title becomes the first Heading 1 and the result says where it came from', async () => {
    const parser = DocxParser.createNewDocument('blank');
    expect(dcTitle(parser)).toBe('Blank Document');
    const result = await composeDocument(parser, '# MCP Engine Evaluation 2026\n\nBody text.\n');
    expect(result.success).toBe(true);
    expect(result.titleSetFrom).toBe('Heading 1');
    expect(dcTitle(parser)).toBe('MCP Engine Evaluation 2026');
  });

  it('a title the file already had is never changed', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    await new OperationEngine(ast, parser).executeBatch({ operations: [{ action: 'metadata', operation: 'setProperties', title: 'Chosen Title' }] as never });
    const result = await composeDocument(parser, '# Something Else\n');
    expect(result.titleSetFrom).toBeUndefined();
    expect(dcTitle(parser)).toBe('Chosen Title');
  });

  it('an explicit metadata title later in the same batch as a cover page wins', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const out = await new OperationEngine(ast, parser).executeBatch({
      operations: [
        { action: 'coverPage', title: 'Cover Title' },
        { action: 'metadata', operation: 'setProperties', title: 'Explicit Title' },
      ] as never,
    });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(dcTitle(parser)).toBe('Explicit Title');
  });

  it('a cover page sets the title from the cover', async () => {
    const parser = DocxParser.createNewDocument('report');
    const ast = await parser.getAST();
    await new OperationEngine(ast, parser).executeBatch({ operations: [{ action: 'coverPage', title: 'Annual Report 2026' }] as never });
    expect(dcTitle(parser)).toBe('Annual Report 2026');
  });

  it('the seed paragraph "Blank Document" is not a candidate, so a document with nothing else keeps the placeholder', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    expect(deriveTitleFromContent(ast)).toBeNull();
    expect(applyDocumentIdentity(ast, parser).titleSetFrom).toBeUndefined();
  });

  it('a body built with plain inserts gets its title at save time (applyDocumentIdentity)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const first = ast.sections[0].nodes[0].id;
    await new OperationEngine(ast, parser).executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', id: 'h_review', anchor: first, position: 'after', text: 'Quarterly Review' },
        { action: 'namedStyle', operation: 'apply', target: 'h_review', styleName: 'Heading1' },
      ] as never,
      sequential: true,
    });
    const result = applyDocumentIdentity(ast, parser);
    expect(result.titleSetFrom).toBe('Heading 1');
    expect(dcTitle(parser)).toBe('Quarterly Review');
  });

  it('cuts a very long heading to the limit and collapses white space', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const long = 'Word '.repeat(100);
    await composeDocument(parser, `# ${long}\n`);
    const title = dcTitle(parser)!;
    expect(title.length).toBeLessThanOrEqual(DERIVED_TITLE_MAX);
    expect(title.startsWith('Word Word')).toBe(true);
    expect(title).not.toMatch(/\s{2,}/);
  });
});

describe('the creator follows a configured author, and only replaces the tool default', () => {
  it('replaces ASTra with the session author, and lastModifiedBy follows', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    ast.settings.defaultAuthor = 'Ada Lovelace';
    const result = applyDocumentIdentity(ast, parser);
    expect(result.creatorSet).toBe('Ada Lovelace');
    expect(dcCreator(parser)).toBe('Ada Lovelace');
    expect(coreXml(parser)).toContain('<cp:lastModifiedBy>Ada Lovelace</cp:lastModifiedBy>');
  });

  it('keeps ASTra when no author is configured', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    expect(applyDocumentIdentity(ast, parser).creatorSet).toBeUndefined();
    expect(dcCreator(parser)).toBe('ASTra');
  });

  it('never replaces a creator the file already had', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    await new OperationEngine(ast, parser).executeBatch({ operations: [{ action: 'metadata', operation: 'setProperties', creator: 'Grace Hopper' }] as never });
    ast.settings.defaultAuthor = 'Ada Lovelace';
    expect(applyDocumentIdentity(ast, parser).creatorSet).toBeUndefined();
    expect(dcCreator(parser)).toBe('Grace Hopper');
  });
});
