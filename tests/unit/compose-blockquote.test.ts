import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, convertAstToMarkdown, validateDocxBuffer } from '@attentionform/astra-engine';
import { compileMarkdown, compileMarkdownDetailed, composeDocument } from '../../packages/engine/src/compose/markdown';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

function part(parser: DocxParser, name: string): string {
  const zip = new AdmZip(parser.rezip());
  return zip.getEntry(name)?.getData().toString('utf8') ?? '';
}

function paragraphsWithStyle(xml: string, style: string): string[] {
  return (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).filter((p) => p.includes(`<w:pStyle w:val="${style}"/>`));
}

describe('plain blockquotes compose to Quote paragraphs', () => {
  it('writes a single quote in the Quote style with no warning', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const md = '# Title\n\n> "Simplicity is prerequisite for reliability." — Edsger Dijkstra\n\nAfter.\n';
    expect(compileMarkdownDetailed(md).warnings).toEqual([]);
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);
    const doc = part(parser, 'word/document.xml');
    const quotes = paragraphsWithStyle(doc, 'Quote');
    expect(quotes).toHaveLength(1);
    expect(quotes[0]).toContain('Simplicity is prerequisite for reliability.');
    expect(quotes[0]).toContain('Edsger Dijkstra');
    // Level 1 takes its indent from the style, not from the paragraph.
    expect(quotes[0]).not.toContain('<w:ind');
    const styles = part(parser, 'word/styles.xml');
    expect(styles).toContain('w:styleId="Quote"');
    expect(styles).toMatch(/w:styleId="Quote"[\s\S]*?<w:pBdr>[\s\S]*?<w:left [^>]*\/>[\s\S]*?<w:i\/>/);
    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('keeps blank quote lines as separate paragraphs of one block', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '> First paragraph.\n>\n> Second paragraph with **bold**.\n');
    expect(result.success).toBe(true);
    const quotes = paragraphsWithStyle(part(parser, 'word/document.xml'), 'Quote');
    expect(quotes).toHaveLength(2);
    expect(quotes[1]).toContain('<w:b/>');
  });

  it('indents nested quotes one step per level and warns past level 3', async () => {
    const md = '> one\n>\n> > two\n> >\n> > > three\n> > >\n> > > > four\n';
    const detailed = compileMarkdownDetailed(md);
    expect(detailed.warnings.some((w) => /deeper than 3 levels/.test(w))).toBe(true);
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);
    const quotes = paragraphsWithStyle(part(parser, 'word/document.xml'), 'Quote');
    expect(quotes).toHaveLength(4);
    const lefts = quotes.map((q) => /<w:ind [^>]*w:left="(\d+)"/.exec(q)?.[1] ?? 'style');
    expect(lefts).toEqual(['style', '1440', '2160', '2160']);
  });

  it('keeps flat lists and fenced code inside a quote', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const md = '> Intro line.\n>\n> - alpha\n> - beta\n>\n> ```js\n> const x = 1;\n> ```\n';
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);
    const doc = part(parser, 'word/document.xml');
    const quotes = paragraphsWithStyle(doc, 'Quote');
    // intro + two list items
    expect(quotes).toHaveLength(3);
    expect(quotes[1]).toContain('<w:numPr>');
    expect(doc).toContain('const x = 1;');
    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('refuses tables, headings and images inside a quote, and unknown alert tags', () => {
    expect(() => compileMarkdown('> | a | b |\n> | - | - |\n> | 1 | 2 |\n')).toThrow(/tables inside a quote/);
    expect(() => compileMarkdown('> # Heading\n')).toThrow(/headings inside a quote/);
    expect(() => compileMarkdown('> ![alt](x.png)\n')).toThrow(/images are out of scope in a quote/);
    expect(() => compileMarkdown('> [!BOGUS]\n> x\n')).toThrow(/unknown alert type \[!BOGUS\]/);
    expect(() => compileMarkdown('>\n')).toThrow(/blockquote is empty/);
  });

  it('reports the line of a refused block inside a quote', () => {
    let message = '';
    try {
      compileMarkdown('Intro.\n\n> fine\n>\n> | a | b |\n> | - | - |\n> | 1 | 2 |\n');
    } catch (err: unknown) {
      message = err instanceof Error ? err.message : String(err);
    }
    expect(message).toMatch(/line 5:/);
  });

  it('adds the Quote style to an opened document that lacks it', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const dom = parser.getStylesDom()!;
    const quote = Array.from(dom.getElementsByTagName('w:style')).find((s) => s.getAttribute('w:styleId') === 'Quote')!;
    quote.parentNode!.removeChild(quote);
    expect(parser.getStyleIds()).not.toContain('Quote');
    const result = await composeDocument(parser, '> quoted\n');
    expect(result.success).toBe(true);
    expect(parser.getStyleIds()).toContain('Quote');
    expect(part(parser, 'word/styles.xml')).toContain('w:styleId="Quote"');
  });

  it('exports Quote paragraphs back to > markers with their nesting', async () => {
    const parser = DocxParser.createNewDocument('blank');
    await composeDocument(parser, '> outer\n>\n> > inner\n\nplain\n');
    const md = convertAstToMarkdown(await parser.getAST(), parser);
    expect(md).toMatch(/^> .*outer/m);
    expect(md).toMatch(/^> > .*inner/m);
    expect(md).toMatch(/^plain$/m);
  });
});
