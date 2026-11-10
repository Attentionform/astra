import { describe, it, expect } from 'vitest';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { executeCoverPage } from '../src/operations/coverPage';

const RPR_ORDER = ['w:rFonts', 'w:b', 'w:color', 'w:sz'];

async function fresh() {
  const parser = new DocxParser(createBlankDocxBuffer('Cover Quality'));
  const ast = await parser.getAST();
  return { parser, ast };
}

describe('cover page is built in schema order', () => {
  for (const template of ['modern_minimalist', 'corporate_navy', 'academic', 'executive'] as const) {
    it(`${template}: every w:rPr lists its children in the order the schema requires`, async () => {
      const { parser, ast } = await fresh();
      executeCoverPage({ action: 'coverPage', template, title: 'T', subtitle: 'S', author: 'A', organization: 'O' }, ast, parser);
      const cover = ast.sections[0].nodes.filter((n) => n.id.startsWith('cover_'));
      expect(cover.length).toBeGreaterThan(0);
      let checked = 0;
      for (const node of cover) {
        const rprs = node.xml.getElementsByTagName('w:rPr');
        for (let i = 0; i < rprs.length; i++) {
          const names: string[] = [];
          for (let c = 0; c < rprs[i].childNodes.length; c++) {
            const child = rprs[i].childNodes[c] as Element;
            if (child.nodeType === 1) names.push(child.tagName);
          }
          const positions = names.map((n) => RPR_ORDER.indexOf(n));
          expect(positions.every((p) => p >= 0), `unexpected child in ${names.join(',')}`).toBe(true);
          expect([...positions].sort((a, b) => a - b), `${template}: ${names.join(',')}`).toEqual(positions);
          checked += 1;
        }
      }
      expect(checked).toBeGreaterThanOrEqual(3);
    });
  }
});
