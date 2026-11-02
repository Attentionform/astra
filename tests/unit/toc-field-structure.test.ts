import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, composeDocument, createBlankDocxBuffer } from '@attentionform/astra-engine';

/**
 * The table of contents is one complex field (begin, instruction, separate, entries, end). It used to be an
 * empty w:fldSimple in the "Contents" paragraph with the entries outside it. Word updates the fields when it
 * opens the file (updateFields), and it then built a second table of contents next to the first.
 */
async function withToc(levels?: number) {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  await composeDocument(parser, '# First\n\n## Second\n\n# Third\n');
  const ast2 = await parser.getAST();
  const engine2 = new OperationEngine(ast2, parser);
  const toc = await engine2.executeBatch({
    operations: [{ action: 'field', operation: 'insertTOC', ...(levels ? { levels } : {}) }] as never,
  });
  expect(toc.success).toBe(true);
  const xml = parser.readDocumentXml();
  const nodes = ast2.sections.flatMap((s) => s.nodes).filter((n) => n.id.startsWith('toc_'));
  return { xml, nodes };
}

const count = (text: string, needle: string): number => text.split(needle).length - 1;

describe('table of contents field', () => {
  it('is one complex field: begin, instruction, separate, entries, end', async () => {
    const { xml } = await withToc();
    expect(xml).not.toContain('w:fldSimple');
    expect(count(xml, 'w:fldCharType="begin"')).toBe(1);
    expect(count(xml, 'w:fldCharType="separate"')).toBe(1);
    expect(count(xml, 'w:fldCharType="end"')).toBe(1);
    expect(xml).toContain('<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>');
    const at = (s: string) => xml.indexOf(s);
    expect(at('w:fldCharType="begin"')).toBeLessThan(at('w:fldCharType="separate"'));
    expect(at('w:fldCharType="separate"')).toBeLessThan(at('w:hyperlink'));
    expect(at('w:hyperlink')).toBeLessThan(at('w:fldCharType="end"'));
  });

  it('starts in the first entry and ends in the last, not in the Contents paragraph', async () => {
    const { nodes } = await withToc();
    const parts = nodes.map((n) => new XMLSerializer().serializeToString(n.xml));
    expect(parts.length).toBeGreaterThanOrEqual(2);
    expect(parts[0]).toContain('Table of Contents');
    expect(parts[0]).not.toContain('fldChar');
    expect(parts[0]).not.toContain('instrText');
    expect(parts[1]).toContain('w:fldCharType="begin"');
    expect(parts[1]).toContain('w:fldCharType="separate"');
    expect(parts[parts.length - 1]).toContain('w:fldCharType="end"');
    for (const middle of parts.slice(2, -1)) expect(middle).not.toContain('fldChar');
  });

  it('every entry paragraph is still a TOC paragraph with a link, a dot leader and a page number', async () => {
    const { nodes } = await withToc();
    const entries = nodes.slice(1);
    expect(entries.length).toBeGreaterThanOrEqual(3);
    for (const e of entries) {
      const x = new XMLSerializer().serializeToString(e.xml);
      expect(x).toMatch(/w:pStyle w:val="TOC[123]"/);
      expect(x).toContain('w:leader="dot"');
      expect(x).toContain('<w:hyperlink');
    }
  });

  it('writes the levels asked for into the instruction', async () => {
    const { xml } = await withToc(2);
    expect(xml).toContain('TOC \\o "1-2" \\h \\z \\u');
  });

  it('keeps asking Word to update the fields on open', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    await engine.executeBatch({ operations: [{ action: 'field', operation: 'insertTOC' }] as never });
    expect(parser.getZipAdapter().readTextOrNull('word/settings.xml')).toContain('<w:updateFields w:val="true"/>');
  });

  it('the metadata of the nodes says which paragraphs hold field characters', async () => {
    const { nodes } = await withToc();
    expect(nodes[0].metadata.hasField).toBe(false);
    expect(nodes[1].metadata.hasField).toBe(true);
    expect(nodes[nodes.length - 1].metadata.hasField).toBe(true);
  });
});
