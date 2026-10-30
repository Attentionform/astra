import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, createBlankDocxBuffer, latexToOmml } from '@attentionform/astra-engine';
import AdmZip from 'adm-zip';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

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

describe('replace and style on a paragraph with inline math', () => {
  const sentence = (p: Element, doc: Document) => {
    p.appendChild(textRun(doc, 'Let '));
    p.appendChild(latexToOmml(doc, '\\frac{a}{b}'));
    p.appendChild(textRun(doc, ' be it.'));
  };
  const mathCount = (parser: DocxParser) => (new AdmZip(parser.rezip()).readAsText('word/document.xml').match(/<m:oMath[ >]/g) ?? []).length;

  it('replace is refused with E_VALIDATION and nothing changes', async () => {
    const { parser, ast, node } = await reloadWith(sentence);
    const before = parser.rezip();
    const res: any = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'replace', target: node.id, text: 'New text' }] as never, rollbackOnError: true,
    });
    expect(res.success).toBe(false);
    expect(String(res.errors[0].error)).toContain('E_VALIDATION');
    expect(String(res.errors[0].error)).toContain('contains math');
    expect(Buffer.compare(parser.rezip(), before)).toBe(0);
  });

  it('replace on a plain paragraph still works', async () => {
    const { parser, ast, node } = await reloadWith((p, doc) => { p.appendChild(textRun(doc, 'Plain text.')); });
    const res: any = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'replace', target: node.id, text: 'New text' }] as never, rollbackOnError: true,
    });
    expect(res.success).toBe(true);
  });

  it('style with match formats the text and keeps the formula', async () => {
    const { parser, ast, node } = await reloadWith(sentence);
    const res: any = await new OperationEngine(ast, parser).executeBatch({
      operations: [{ action: 'style', target: node.id, bold: true, match: 'Let' }] as never, rollbackOnError: true,
    });
    expect(res.success).toBe(true);
    expect(mathCount(parser)).toBe(1);
  });
});
