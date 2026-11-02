import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
import {
  compileMarkdown,
  composeDocument,
} from '../../packages/engine/src/compose/markdown';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

function documentXml(parser: DocxParser): string {
  const zip = new AdmZip(parser.rezip());
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

function parseParagraphElements(xml: string): Element[] {
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const paras = Array.from(doc.getElementsByTagName('w:p'));
  return paras;
}

describe('Gap 9: footnote and endnote reference placement', () => {
  it('places three footnote references at their exact inline positions', async () => {
    const md = 'First[^a] middle[^b] end[^c].\n\n[^a]: Note A\n\n[^b]: Note B\n\n[^c]: Note C\n';
    const ops = compileMarkdown(md);

    const fnOps = ops.filter((op) => op.action === 'footnote');
    expect(fnOps).toHaveLength(3);
    expect((fnOps[0] as { offset?: number }).offset).toBe(5);
    expect((fnOps[1] as { offset?: number }).offset).toBe(12);
    expect((fnOps[2] as { offset?: number }).offset).toBe(16);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('First'));
    expect(p).toBeDefined();

    const children = Array.from(p!.childNodes).filter((n) => n.nodeType === 1) as Element[];
    // Examine non-pPr children
    const contentNodes = children.filter((el) => el.tagName !== 'w:pPr');

    // Expected sequence:
    // 1. run with text "First"
    // 2. run with footnoteReference id="1"
    // 3. run with text " middle"
    // 4. run with footnoteReference id="2"
    // 5. run with text " end"
    // 6. run with footnoteReference id="3"
    // 7. run with text "."
    expect(contentNodes).toHaveLength(7);

    expect(contentNodes[0].textContent).toBe('First');
    expect(contentNodes[1].getElementsByTagName('w:footnoteReference')).toHaveLength(1);
    expect(contentNodes[1].getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('1');
    expect(contentNodes[2].textContent).toBe(' middle');
    expect(contentNodes[3].getElementsByTagName('w:footnoteReference')).toHaveLength(1);
    expect(contentNodes[3].getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('2');
    expect(contentNodes[4].textContent).toBe(' end');
    expect(contentNodes[5].getElementsByTagName('w:footnoteReference')).toHaveLength(1);
    expect(contentNodes[5].getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('3');
    expect(contentNodes[6].textContent).toBe('.');

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('preserves bold and italic formatting on both sides of a split run', async () => {
    const md = 'Start **bold[^a] word** and *italic[^b] text* end.\n\n[^a]: Bold note\n\n[^b]: Italic note\n';
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('Start'));
    expect(p).toBeDefined();

    const runs = Array.from(p!.getElementsByTagName('w:r'));
    // Find bold footnote ref
    const fn1RunIdx = runs.findIndex((r) => r.getElementsByTagName('w:footnoteReference').length > 0 && r.getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id') === '1');
    expect(fn1RunIdx).toBeGreaterThan(0);

    const beforeFn1 = runs[fn1RunIdx - 1];
    const afterFn1 = runs[fn1RunIdx + 1];

    expect(beforeFn1.textContent).toBe('bold');
    expect(beforeFn1.getElementsByTagName('w:b')).toHaveLength(1);

    expect(afterFn1.textContent).toBe(' word');
    expect(afterFn1.getElementsByTagName('w:b')).toHaveLength(1);

    // Find italic footnote ref
    const fn2RunIdx = runs.findIndex((r) => r.getElementsByTagName('w:footnoteReference').length > 0 && r.getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id') === '2');
    expect(fn2RunIdx).toBeGreaterThan(0);

    const beforeFn2 = runs[fn2RunIdx - 1];
    const afterFn2 = runs[fn2RunIdx + 1];

    expect(beforeFn2.textContent).toBe('italic');
    expect(beforeFn2.getElementsByTagName('w:i')).toHaveLength(1);

    expect(afterFn2.textContent).toBe(' text');
    expect(afterFn2.getElementsByTagName('w:i')).toHaveLength(1);

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('supports two adjacent footnotes in the same run and keeps document order', async () => {
    const md = 'Adjacent[^a][^b] notes.\n\n[^a]: First\n\n[^b]: Second\n';
    const ops = compileMarkdown(md);
    const fnOps = ops.filter((op) => op.action === 'footnote');
    expect((fnOps[0] as { offset?: number }).offset).toBe(8);
    expect((fnOps[1] as { offset?: number }).offset).toBe(8);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('Adjacent'));
    expect(p).toBeDefined();

    const contentNodes = Array.from(p!.childNodes).filter((n) => n.nodeType === 1 && (n as Element).tagName !== 'w:pPr') as Element[];
    expect(contentNodes).toHaveLength(4);
    expect(contentNodes[0].textContent).toBe('Adjacent');
    expect(contentNodes[1].getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('1');
    expect(contentNodes[2].getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('2');
    expect(contentNodes[3].textContent).toBe(' notes.');

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
  });

  it('handles footnote reference at very start and very end of a paragraph', async () => {
    const md = '[^start]Beginning and end.[^end]\n\n[^start]: Note start\n\n[^end]: Note end\n';
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('Beginning'));
    expect(p).toBeDefined();

    const contentNodes = Array.from(p!.childNodes).filter((n) => n.nodeType === 1 && (n as Element).tagName !== 'w:pPr') as Element[];
    expect(contentNodes[0].getElementsByTagName('w:footnoteReference')).toHaveLength(1);
    expect(contentNodes[0].getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('1');

    const last = contentNodes[contentNodes.length - 1];
    expect(last.getElementsByTagName('w:footnoteReference')).toHaveLength(1);
    expect(last.getElementsByTagName('w:footnoteReference')[0].getAttribute('w:id')).toBe('2');

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
  });

  it('places footnote reference directly after a hyperlink as a sibling element', async () => {
    const md = 'Visit [our site](https://example.com)[^link] for details.\n\n[^link]: Link note\n';
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('Visit'));
    expect(p).toBeDefined();

    const children = Array.from(p!.childNodes).filter((n) => n.nodeType === 1 && (n as Element).tagName !== 'w:pPr') as Element[];
    const linkIdx = children.findIndex((el) => el.tagName === 'w:hyperlink');
    expect(linkIdx).toBeGreaterThan(0);

    // Footnote run should be right after w:hyperlink as a direct child of w:p
    const afterLink = children[linkIdx + 1];
    expect(afterLink.tagName).toBe('w:r');
    expect(afterLink.getElementsByTagName('w:footnoteReference')).toHaveLength(1);

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
  });

  it('correctly offsets footnote reference when adjacent to a citation', async () => {
    const md = 'According to [@doe24][^fn] earlier findings.\n\n[@doe24]: Doe, 2024\n\n[^fn]: Check chapter 2\n';
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('According to'));
    expect(p).toBeDefined();

    // Citation [1] should precede the footnote reference
    const runs = Array.from(p!.getElementsByTagName('w:r'));
    const fnIdx = runs.findIndex((r) => r.getElementsByTagName('w:footnoteReference').length > 0);
    expect(fnIdx).toBeGreaterThan(0);

    const beforeFn = runs[fnIdx - 1];
    expect(beforeFn.textContent).toContain('[1]');

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
  });

  it('correctly places endnote reference at typed offset', async () => {
    const md = 'Endnote[^en-1] here.\n\n[^en-1]: The end body\n';
    const ops = compileMarkdown(md);
    const endOp = ops.find((op) => op.action === 'endnote');
    expect(endOp).toBeDefined();
    expect((endOp as { offset?: number }).offset).toBe(7);

    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const xml = documentXml(parser);
    const paras = parseParagraphElements(xml);
    const p = paras.find((node) => (node.textContent || '').includes('Endnote'));
    expect(p).toBeDefined();

    const contentNodes = Array.from(p!.childNodes).filter((n) => n.nodeType === 1 && (n as Element).tagName !== 'w:pPr') as Element[];
    expect(contentNodes).toHaveLength(3);
    expect(contentNodes[0].textContent).toBe('Endnote');
    expect(contentNodes[1].getElementsByTagName('w:endnoteReference')).toHaveLength(1);
    expect(contentNodes[2].textContent).toBe(' here.');

    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
  });
});
