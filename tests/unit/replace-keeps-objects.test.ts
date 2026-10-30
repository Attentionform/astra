import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, executeReplace, assertValidDocx } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const GALLERY = path.resolve(__dirname, '../../test-gallery');

function thrown(fn: () => unknown): { code?: string; message: string } {
  try {
    fn();
  } catch (e) {
    const err = e as { code?: string; message?: string };
    return { code: err.code, message: String(err.message) };
  }
  throw new Error('expected the call to throw');
}

/** T5_image_doc.docx with text added to its picture paragraph, as a separate run or inside the picture's own run. */
async function loadParagraphWithPicture(sameRun: boolean) {
  const p1 = new DocxParser(fs.readFileSync(path.join(GALLERY, 'T5_image_doc.docx')));
  const ast1 = await p1.getAST();
  const imgNode = ast1.sections.flatMap((s) => s.nodes).find((n) => n.type === 'image')!;
  const xml = imgNode.xml as Element;
  const t = xml.ownerDocument.createElement('w:t');
  t.setAttribute('xml:space', 'preserve');
  t.textContent = 'Text beside the picture. ';
  if (sameRun) {
    const drawingRun = xml.getElementsByTagName('w:drawing')[0].parentNode as Element;
    drawingRun.insertBefore(t, drawingRun.firstChild);
  } else {
    const run = xml.ownerDocument.createElement('w:r');
    run.appendChild(t);
    const pPr = xml.getElementsByTagName('w:pPr')[0];
    xml.insertBefore(run, pPr ? pPr.nextSibling : xml.firstChild);
  }
  const parser = new DocxParser(p1.rezip());
  const ast = await parser.getAST();
  const node = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === imgNode.id)!;
  return { parser, ast, node };
}

describe('replace keeps pictures and comment anchors, and refuses object nodes', () => {
  it.each([[false], [true]])('untracked replace keeps the picture (text in the picture run: %s)', async (sameRun) => {
    const { ast, node } = await loadParagraphWithPicture(sameRun);
    expect(node.type).toBe('paragraph');
    const res = executeReplace({ action: 'replace', target: node.id, text: 'New text only', trackChanges: false }, ast);
    const el = node.xml as Element;
    expect(el.getElementsByTagName('w:drawing').length).toBe(1);
    expect(el.textContent).toContain('New text only');
    expect(el.textContent).not.toContain('Text beside the picture');
    expect(res.warnings.join(' ')).toContain('moved them to the end of the paragraph');
  });

  it('tracked replace keeps the picture and the deleted text', async () => {
    const { parser, ast, node } = await loadParagraphWithPicture(false);
    executeReplace({ action: 'replace', target: node.id, text: 'New text only', trackChanges: true }, ast);
    const el = node.xml as Element;
    expect(el.getElementsByTagName('w:drawing').length).toBe(1);
    expect(el.getElementsByTagName('w:del').length).toBeGreaterThan(0);
    await assertValidDocx(parser.rezip());
  });

  it('replace on an image node is refused and leaves the picture', async () => {
    const parser = new DocxParser(fs.readFileSync(path.join(GALLERY, 'T5_image_doc.docx')));
    const ast = await parser.getAST();
    const img = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'image')!;
    const before = parser.rezip();
    const err = thrown(() => executeReplace({ action: 'replace', target: img.id, text: 'hello' }, ast));
    expect(err.code).toBe('E_VALIDATION');
    expect(err.message).toContain('is a image node');
    expect(Buffer.compare(parser.rezip(), before)).toBe(0);
  });

  it('replace on an equation node and on a chart node is refused', async () => {
    const parser = new DocxParser(fs.readFileSync(path.join(GALLERY, 'T2_edit_target.docx')));
    const ast = await parser.getAST();
    const first = ast.sections[0].nodes[0].id;
    const engine = new OperationEngine(ast, parser);
    const r = await engine.executeBatch({
      rollbackOnError: true,
      operations: [
        { action: 'equation', operation: 'insert', anchor: first, position: 'after', latex: '\\frac{a}{b}', id: 'eq_1' },
        { action: 'insert', type: 'chart', anchor: 'eq_1', position: 'after', chartType: 'column', title: 'T',
          categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }], id: 'chart_1' },
      ] as never,
    });
    expect(r.success).toBe(true);
    const reloaded = new DocxParser(parser.rezip());
    const ast2 = await reloaded.getAST();
    const nodes = ast2.sections.flatMap((s) => s.nodes);
    for (const type of ['equation', 'chart']) {
      const node = nodes.find((n) => n.type === type)!;
      const err = thrown(() => executeReplace({ action: 'replace', target: node.id, text: 'x' }, ast2));
      expect(err.code).toBe('E_VALIDATION');
      expect(err.message).toContain(`is a ${type} node`);
    }
  });

  it('replace on a commented paragraph keeps the comment reference', async () => {
    const parser = new DocxParser(fs.readFileSync(path.join(GALLERY, 'T2_edit_target.docx')));
    const ast = await parser.getAST();
    const target = ast.sections[0].nodes.find((n) => n.type === 'paragraph' && (n.metadata.text ?? '').length > 0)!;
    const engine = new OperationEngine(ast, parser);
    const r = await engine.executeBatch({
      rollbackOnError: true,
      operations: [{ action: 'comment', operation: 'insert', target: target.id, text: 'check' }] as never,
    });
    expect(r.success).toBe(true);
    executeReplace({ action: 'replace', target: target.id, text: 'Brand new text' }, ast);
    const docXml = new AdmZip(parser.rezip()).readAsText('word/document.xml');
    expect(docXml.match(/<w:commentReference /g)?.length).toBe(1);
    expect(docXml).toContain('Brand new text');
  });
});
