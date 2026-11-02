import { describe, it, expect } from 'vitest';
import { DocxParser, convertAstToHtml, htmlDocumentTitle, type DocumentAST } from '@attentionform/astra-engine';
import { handleExportDocument } from '../../packages/mcp-server/src/tools/export';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';

/**
 * The `<title>` of an exported page is the document's own title when it has one, else the file name.
 * A new document carries its template name as the title (`Blank Document`), which is a placeholder.
 */
const astWith = (title: unknown): DocumentAST => ({ metadata: { title } }) as unknown as DocumentAST;

describe('htmlDocumentTitle', () => {
  it('uses the document title', () => {
    expect(htmlDocumentTitle(astWith('Quarterly Review'), 'q3')).toBe('Quarterly Review');
  });

  it('falls back to the file name when there is no title, an empty one or one of white space', () => {
    expect(htmlDocumentTitle(astWith(undefined), 'q3')).toBe('q3');
    expect(htmlDocumentTitle(astWith(''), 'q3')).toBe('q3');
    expect(htmlDocumentTitle(astWith('  \n\t '), 'q3')).toBe('q3');
    expect(htmlDocumentTitle({} as unknown as DocumentAST, 'q3')).toBe('q3');
  });

  it('joins line breaks and runs of white space into single spaces', () => {
    expect(htmlDocumentTitle(astWith('  Annual\n   report\t2026 '), 'q3')).toBe('Annual report 2026');
  });

  it('treats the name of a template as no title', () => {
    for (const placeholder of ['Blank Document', 'Report', 'Memo', 'Api Spec']) {
      expect(htmlDocumentTitle(astWith(placeholder), 'q3')).toBe('q3');
    }
    expect(htmlDocumentTitle(astWith('Report on cloud costs'), 'q3')).toBe('Report on cloud costs');
  });

  it('a title that is not a string is ignored', () => {
    expect(htmlDocumentTitle(astWith(42), 'q3')).toBe('q3');
  });
});

describe('page title of an exported document', () => {
  async function titled(title: string | null): Promise<DocxParser> {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    if (title !== null) {
      const { OperationEngine } = await import('@attentionform/astra-engine');
      const result = await new OperationEngine(ast, parser).executeBatch({
        operations: [{ action: 'metadata', operation: 'setProperties', title }] as never,
        sequential: true,
      });
      expect(result.success).toBe(true);
    }
    return parser;
  }

  it('shows the title that was set, with markup escaped', async () => {
    const parser = await titled('Quarterly <Review> & Plan');
    const ast = await parser.getAST();
    const html = convertAstToHtml(ast, htmlDocumentTitle(ast, 'file-name'), parser);
    expect(html).toContain('<title>Quarterly &lt;Review&gt; &amp; Plan</title>');
  });

  it('shows the file name for a new document that was given no title', async () => {
    const parser = await titled(null);
    const ast = await parser.getAST();
    const html = convertAstToHtml(ast, htmlDocumentTitle(ast, 'file-name'), parser);
    expect(html).toContain('<title>file-name</title>');
  });

  it('MCP export_document uses the title of the live session', async () => {
    const parser = await titled('Board pack');
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(parser.rezip(), 'deck.docx');
    const out = await handleExportDocument(store, { sessionId: handle.sessionId, format: 'html' });
    expect(out.ok).toBe(true);
    if (out.ok) expect(JSON.stringify(out)).toContain('<title>Board pack</title>');
  });
});
