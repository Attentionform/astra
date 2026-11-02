import { describe, it, expect } from 'vitest';
import { DocxParser, composeDocument, compileMarkdown } from '@attentionform/astra-engine';

/**
 * Two tables with no paragraph between them touch in Word. A callout is a one-cell table, so a callout after a
 * table, a table after a callout and two callouts in a row get an empty paragraph between them when a document
 * is composed from Markdown.
 */
const TABLE = '| A | B |\n| --- | --- |\n| 1 | 2 |';
const CALLOUT = '> [!WARNING]\n> Careful here.';

async function composed(md: string): Promise<string> {
  const parser = DocxParser.createNewDocument('blank');
  const res = await composeDocument(parser, md);
  expect(res.success).toBe(true);
  return parser.readDocumentXml();
}

const adjacentTables = (xml: string): boolean => /<\/w:tbl>\s*<w:tbl[ >]/.test(xml);

describe('empty paragraph between neighbouring tables', () => {
  it('a callout after a table', async () => {
    const xml = await composed(`${TABLE}\n\n${CALLOUT}\n\nafter`);
    expect(adjacentTables(xml)).toBe(false);
  });

  it('a table after a callout', async () => {
    const xml = await composed(`${CALLOUT}\n\n${TABLE}\n\nafter`);
    expect(adjacentTables(xml)).toBe(false);
  });

  it('two callouts in a row', async () => {
    const xml = await composed(`${CALLOUT}\n\n> [!NOTE]\n> Second.\n\nafter`);
    expect(adjacentTables(xml)).toBe(false);
  });

  it('a table after a table', async () => {
    const xml = await composed(`${TABLE}\n\n${TABLE}\n\nafter`);
    expect(adjacentTables(xml)).toBe(false);
  });

  it('a paragraph before a callout gets no extra paragraph', () => {
    const ops = compileMarkdown(`Intro text\n\n${CALLOUT}`);
    const empty = ops.filter((o) => o.action === 'insert' && (o as { type?: string }).type === 'paragraph' && (o as { text?: string }).text === '');
    expect(empty.length).toBe(0);
  });

  it('the paragraph that is added is empty and sits before the callout', () => {
    const ops = compileMarkdown(`${TABLE}\n\n${CALLOUT}`);
    const at = ops.findIndex((o) => o.action === 'insert' && (o as { type?: string }).type === 'paragraph' && (o as { text?: string }).text === '');
    const callout = ops.findIndex((o) => o.action === 'callout');
    expect(at).toBeGreaterThan(-1);
    expect(at).toBeLessThan(callout);
  });

  it('every table keeps its own content', async () => {
    const xml = await composed(`${TABLE}\n\n${CALLOUT}\n\nafter`);
    expect(xml).toContain('Careful here.');
    expect((xml.match(/<w:tbl>/g) ?? []).length).toBe(2);
  });
});
