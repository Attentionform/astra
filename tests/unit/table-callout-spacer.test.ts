import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, composeDocument, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

/**
 * Two tables with no paragraph between them touch in Word (LibreOffice leaves a gap). A callout is a one-cell
 * table. When an operation adds a table or a callout right next to a table, the engine puts an empty paragraph
 * between them, whether the document is built by operations (batch, MCP execute_batch) or composed from Markdown.
 * The paragraph is an ordinary one: a caller can restyle or delete it.
 */
const TABLE = (id: string, anchor = 'end') =>
  ({ action: 'insert', type: 'table', anchor, headers: ['A', 'B'], rows: [['1', '2']], id }) as const;
const CALLOUT = (anchor: string, position: 'after' | 'before' = 'after') =>
  ({ action: 'callout', anchor, position, text: 'Careful here.', theme: 'amber' }) as const;

async function build(operations: unknown[]) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  // A blank document starts with its own paragraphs; the checks look at what the operations add after them.
  const skip = ast.sections.flatMap((sec) => sec.nodes).length;
  const engine = new OperationEngine(ast, parser);
  const res = await engine.executeBatch({ operations: operations as never, sequential: true });
  expect(res.success).toBe(true);
  return { parser, ast, engine, skip, xml: () => parser.readDocumentXml() };
}

const adjacentTables = (xml: string): boolean => /<\/w:tbl>\s*<w:tbl[ >]/.test(xml);
const added = (built: Awaited<ReturnType<typeof build>>) => built.ast.sections.flatMap((s) => s.nodes).slice(built.skip);
const types = (built: Awaited<ReturnType<typeof build>>): string[] => added(built).map((n) => n.type);

describe('operations that add a table next to a table', () => {
  it('a callout after a table', async () => {
    const built = await build([TABLE('t1'), CALLOUT('t1')]);
    expect(adjacentTables(built.xml())).toBe(false);
    expect(types(built)).toEqual(['table', 'paragraph', 'table']);
  });

  it('a callout before a table', async () => {
    const { xml } = await build([TABLE('t1'), CALLOUT('t1', 'before')]);
    expect(adjacentTables(xml())).toBe(false);
  });

  it('a callout between two tables gets a paragraph on each side', async () => {
    const built = await build([TABLE('t1'), TABLE('t2'), CALLOUT('t1')]);
    expect(adjacentTables(built.xml())).toBe(false);
    expect(types(built)).toEqual(['table', 'paragraph', 'table', 'paragraph', 'table']);
  });

  it('a table after a table', async () => {
    const { xml } = await build([TABLE('t1'), TABLE('t2', 't1')]);
    expect(adjacentTables(xml())).toBe(false);
  });

  it('two callouts in a row', async () => {
    const { xml } = await build([CALLOUT('end'), CALLOUT('end')]);
    expect(adjacentTables(xml())).toBe(false);
  });

  it('a callout after a paragraph gets no extra paragraph', async () => {
    const built = await build([
      { action: 'insert', type: 'paragraph', anchor: 'end', text: 'Intro', id: 'p1' },
      CALLOUT('p1'),
    ]);
    expect(types(built)).toEqual(['paragraph', 'table']);
  });

  it('a table after a paragraph gets no extra paragraph', async () => {
    const built = await build([
      { action: 'insert', type: 'paragraph', anchor: 'end', text: 'Intro', id: 'p1' },
      TABLE('t1', 'p1'),
    ]);
    expect(types(built)).toEqual(['paragraph', 'table']);
  });

  it('tables that were far apart and stay far apart are left alone', async () => {
    const built = await build([
      TABLE('t1'),
      { action: 'insert', type: 'paragraph', anchor: 't1', text: 'between', id: 'p1' },
      TABLE('t2', 'p1'),
    ]);
    expect(types(built)).toEqual(['table', 'paragraph', 'table']);
  });

  it('the paragraph that is added is empty', async () => {
    const built = await build([TABLE('t1'), CALLOUT('t1')]);
    const gap = added(built)[1];
    expect(gap.type).toBe('paragraph');
    expect((gap.metadata as { text?: string }).text ?? '').toBe('');
  });

  it('every table keeps its own content', async () => {
    const { xml } = await build([TABLE('t1'), CALLOUT('t1')]);
    expect(xml()).toContain('Careful here.');
    expect((xml().match(/<w:tbl>/g) ?? []).length).toBe(2);
  });
});

describe('the paragraph is under the caller\'s control', () => {
  it('can be restyled with a style operation', async () => {
    const built = await build([TABLE('t1'), CALLOUT('t1')]);
    const { engine, xml } = built;
    const gap = added(built)[1];
    const res = await engine.executeBatch({ operations: [{ action: 'style', target: gap.id, spacingAfter: 600 }] as never });
    expect(res.success).toBe(true);
    expect(xml()).toContain('w:after="600"');
  });

  it('can be deleted, and the engine does not put it back', async () => {
    const built = await build([TABLE('t1'), CALLOUT('t1')]);
    const { engine, xml } = built;
    const gap = added(built)[1];
    const res = await engine.executeBatch({ operations: [{ action: 'delete', target: gap.id }] as never });
    expect(res.success).toBe(true);
    expect(adjacentTables(xml())).toBe(true);
  });
});

describe('documents composed from Markdown', () => {
  const MD_TABLE = '| A | B |\n| --- | --- |\n| 1 | 2 |';
  const MD_CALLOUT = '> [!WARNING]\n> Careful here.';
  const composed = async (md: string): Promise<string> => {
    const parser = DocxParser.createNewDocument('blank');
    expect((await composeDocument(parser, md)).success).toBe(true);
    return parser.readDocumentXml();
  };

  it('a callout after a table', async () => {
    expect(adjacentTables(await composed(`${MD_TABLE}\n\n${MD_CALLOUT}\n\nafter`))).toBe(false);
  });

  it('a table after a callout', async () => {
    expect(adjacentTables(await composed(`${MD_CALLOUT}\n\n${MD_TABLE}\n\nafter`))).toBe(false);
  });

  it('two callouts in a row', async () => {
    expect(adjacentTables(await composed(`${MD_CALLOUT}\n\n> [!NOTE]\n> Second.\n\nafter`))).toBe(false);
  });

  it('a table after a table', async () => {
    expect(adjacentTables(await composed(`${MD_TABLE}\n\n${MD_TABLE}\n\nafter`))).toBe(false);
  });
});
