import { describe, it, expect } from 'vitest';
import { DocxParser, composeDocument } from '@attentionform/astra-engine';

/**
 * Word draws a paragraph border outside the indent, by the border's w:space (4 pt = 80 twips). A code block's
 * box therefore sits on the text margins when the indent equals the border space on both sides. A left indent of
 * 180 alone put the box 5 pt right of the margin and let it hang 4 pt past the right margin.
 */
describe('code block box', () => {
  it('is indented by the border space on both sides, so it lines up with the text margins', async () => {
    const parser = DocxParser.createNewDocument('blank');
    expect((await composeDocument(parser, '```js\nconsole.log(1);\n```')).success).toBe(true);
    const xml = parser.readDocumentXml();
    expect(xml).toContain('<w:ind w:left="80" w:right="80"/>');
    expect(xml).toContain('w:space="4"');
  });
});
