import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, createBlankDocxBuffer, composeDocument } from '@attentionform/astra-engine';

async function documentXml(md: string): Promise<string> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  await composeDocument(parser, md);
  return new AdmZip(parser.rezip()).getEntry('word/document.xml')!.getData().toString('utf8');
}

describe('composed inline text is escaped exactly once', () => {
  it('keeps apostrophes, ampersands, quotes and angle brackets literal in paragraphs, headings and list items', async () => {
    const xml = await documentXml(
      [
        "# The party's brief",
        '',
        `The party's duty & days' notice, "quoted", 3 < 4 > 2.`,
        '',
        "- item's & more",
        '',
      ].join('\n')
    );
    expect(xml).not.toMatch(/&amp;#39;|&amp;amp;|&amp;lt;|&amp;gt;|&amp;quot;/);
    expect(xml).toContain("The party's brief");
    expect(xml).toContain("The party's duty &amp; days' notice, \"quoted\", 3 &lt; 4 &gt; 2.");
    expect(xml).toContain("item's &amp; more");
  });

  it('covers bold, italic, code and link text', async () => {
    const xml = await documentXml("**it's** *a&b* `x<y` [we're](https://example.com/a?b=1&c=2)\n");
    expect(xml).not.toMatch(/&amp;(?:#39|amp|lt|gt|quot);/);
    expect(xml).toContain("it's");
    expect(xml).toContain('a&amp;b');
    expect(xml).toContain('x&lt;y');
    expect(xml).toContain("we're");
  });

  it('follows CommonMark for authored entities: &lt; and &amp; mean < and &', async () => {
    const xml = await documentXml('Write a &lt; b &amp; c.\n');
    expect(xml).toContain('a &lt; b &amp; c');
    expect(xml).not.toMatch(/&amp;(?:lt|amp);/);
  });

  it('leaves tables, code blocks and callouts as they were', async () => {
    const xml = await documentXml("| a | b |\n|---|---|\n| it's | x & y |\n\n```\nit's <b>\n```\n");
    expect(xml).not.toMatch(/&amp;#39;|&amp;amp;(?!amp)/);
    expect(xml).toContain("it's");
  });
});
