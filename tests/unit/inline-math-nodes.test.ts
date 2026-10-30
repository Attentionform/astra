import { describe, it, expect } from 'vitest';
import { DocxParser, createBlankDocxBuffer, latexToOmml } from '@attentionform/astra-engine';
import { DOMParser } from '@xmldom/xmldom';

const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

/** A blank document whose first paragraph is rebuilt from `build`, saved and reloaded. */
async function reloadWith(build: (p: Element, doc: Document) => void) {
  const seed = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await seed.getAST();
  const p = ast.sections[0].nodes.find((n) => n.xml && (n.xml as Element).tagName === 'w:p')!.xml as Element;
  while (p.firstChild) p.removeChild(p.firstChild);
  build(p, p.ownerDocument as Document);
  const parser = new DocxParser(seed.rezip());
  const reloaded = await parser.getAST();
  return { parser, ast: reloaded, node: reloaded.sections[0].nodes.find((n) => n.xml && (n.xml as Element).tagName === 'w:p')! };
}

function textRun(doc: Document, text: string): Element {
  const r = doc.createElement('w:r');
  const t = doc.createElement('w:t');
  t.setAttribute('xml:space', 'preserve');
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

describe('paragraphs with inline math after a reload', () => {
  it('a sentence with a formula is a paragraph whose text shows the formula as \\(latex\\)', async () => {
    const { node } = await reloadWith((p, doc) => {
      p.appendChild(textRun(doc, 'Let '));
      p.appendChild(latexToOmml(doc, '\\frac{a}{b}'));
      p.appendChild(textRun(doc, ' be it.'));
    });
    expect(node.type).toBe('paragraph');
    expect(node.metadata.hasEquation).toBe(true);
    expect(node.metadata.text).toBe('Let \\(\\frac{a}{b}\\) be it.');
  });

  it('keeps document order with several formulas and plain symbols', async () => {
    const { node } = await reloadWith((p, doc) => {
      p.appendChild(latexToOmml(doc, 'd_k'));
      p.appendChild(textRun(doc, ' and '));
      p.appendChild(latexToOmml(doc, '\\alpha x'));
      p.appendChild(textRun(doc, '.'));
    });
    expect(node.type).toBe('paragraph');
    expect(node.metadata.text).toBe('\\(d_k\\) and \\(\\alpha x\\).');
  });

  it('a paragraph that holds only an inline formula stays an equation node', async () => {
    const { node } = await reloadWith((p, doc) => { p.appendChild(latexToOmml(doc, '\\frac{a}{b}')); });
    expect(node.type).toBe('equation');
    expect(node.metadata.text).toBe('');
  });

  it('a display equation stays an equation node and keeps its empty text', async () => {
    const { node } = await reloadWith((p, doc) => {
      const para = doc.createElement('m:oMathPara');
      para.appendChild(latexToOmml(doc, '\\frac{a}{b}'));
      p.appendChild(para);
    });
    expect(node.type).toBe('equation');
    expect(node.metadata.text).toBe('');
  });

  it('a formula the converter never emits falls back to its flat text', async () => {
    const { node } = await reloadWith((p, doc) => {
      p.appendChild(textRun(doc, 'See '));
      const frag = new DOMParser().parseFromString(
        `<m:oMath ${MATH_NS}><m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>`, 'text/xml').documentElement;
      p.appendChild(doc.importNode(frag as never, true) as unknown as Element);
    });
    expect(node.type).toBe('paragraph');
    expect(node.metadata.text).toBe('See \\(x\\)');
  });

  it('search finds the formula text', async () => {
    const { parser, ast, node } = await reloadWith((p, doc) => {
      p.appendChild(textRun(doc, 'Let '));
      p.appendChild(latexToOmml(doc, '\\frac{a}{b}'));
      p.appendChild(textRun(doc, ' be it.'));
    });
    expect(parser.searchNodesInAST(ast, 'frac').map((n) => n.id)).toContain(node.id);
  });

  it('text followed by a formula at the end is a paragraph', async () => {
    const { node } = await reloadWith((p, doc) => {
      p.appendChild(textRun(doc, 'x '));
      p.appendChild(latexToOmml(doc, 'y'));
    });
    expect(node.type).toBe('paragraph');
  });
});
