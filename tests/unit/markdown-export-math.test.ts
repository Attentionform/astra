import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer, latexToOmml } from '@attentionform/astra-engine';
import { convertAstToMarkdown } from '../../packages/engine/src/export/markdown';

const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

function textRun(doc: Document, text: string): Element {
  const r = doc.createElement('w:r');
  const t = doc.createElement('w:t');
  t.setAttribute('xml:space', 'preserve');
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

/** A blank document whose first paragraph is rebuilt from `build`, saved, reloaded and exported to Markdown. */
async function exportWith(build: (p: Element, doc: Document) => void): Promise<string> {
  const seed = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await seed.getAST();
  const p = ast.sections[0].nodes.find((n) => n.xml && (n.xml as Element).tagName === 'w:p')!.xml as Element;
  while (p.firstChild) p.removeChild(p.firstChild);
  build(p, p.ownerDocument as Document);
  const reloaded = await new DocxParser(seed.rezip()).getAST();
  return String(await convertAstToMarkdown(reloaded as never)).trim();
}

describe('Markdown export of math', () => {
  it('a sentence with a formula exports as \\(latex\\)', async () => {
    const md = await exportWith((p, doc) => {
      p.appendChild(textRun(doc, 'Let '));
      p.appendChild(latexToOmml(doc, '\\frac{a}{b}'));
      p.appendChild(textRun(doc, ' be it.'));
    });
    expect(md).toContain('Let \\(\\frac{a}{b}\\) be it.');
  });

  it('a reloaded display equation exports as $$ LaTeX $$, not the flattened text', async () => {
    const md = await exportWith((p, doc) => {
      const para = doc.createElement('m:oMathPara');
      para.appendChild(latexToOmml(doc, '\\frac{a}{b}'));
      p.appendChild(para);
    });
    expect(md).toContain('$$\n\\frac{a}{b}\n$$');
    expect(md).not.toMatch(/\$\$\s*ab\s*\$\$/);
  });

  it('a paragraph that is only an inline formula exports as a display equation with LaTeX', async () => {
    const md = await exportWith((p, doc) => { p.appendChild(latexToOmml(doc, 'x_i^2')); });
    expect(md).toContain('$$\nx_i^2\n$$');
  });

  it('a formula the converter does not emit falls back to the flat text, wrapped the same way', async () => {
    const md = await exportWith((p, doc) => {
      p.appendChild(textRun(doc, 'See '));
      const frag = new DOMParser().parseFromString(
        `<m:oMath ${MATH_NS}><m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>`, 'text/xml').documentElement;
      p.appendChild(doc.importNode(frag as never, true) as unknown as Element);
    });
    expect(md).toContain('See \\(x\\)');
  });

  it('two formulas in one display paragraph keep all their text', async () => {
    const md = await exportWith((p, doc) => {
      const para = doc.createElement('m:oMathPara');
      para.appendChild(latexToOmml(doc, 'a'));
      para.appendChild(latexToOmml(doc, 'b'));
      p.appendChild(para);
    });
    expect(md).toContain('ab');
  });
});
