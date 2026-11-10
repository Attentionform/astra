import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '@attentionform/astra-engine';
import { compileMarkdown, composeDocument } from '../../packages/engine/src/compose/markdown';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

function documentXml(parser: DocxParser): string {
  const zip = new AdmZip(parser.rezip());
  return zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

describe('inline formatting inside alert callouts', () => {
  it('keeps bold, italic, code and links in the body and the title', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const md = [
      '> [!WARNING]',
      '> **Read *this* first**',
      '> Use **bold**, *italic*, `code` and [a link](https://example.com/x) here.',
      '',
    ].join('\n');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    expect(xml).toContain('FFFBEB');
    // Body: bold run, italic run, code font run and a real hyperlink.
    const tbl = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect(tbl).toMatch(/<w:b\/>[\s\S]*?<w:t[^>]*>bold<\/w:t>/);
    expect(tbl).toMatch(/<w:i\/>[\s\S]*?<w:t[^>]*>italic<\/w:t>/);
    expect(tbl).toContain('Courier New');
    expect(tbl).toContain('<w:hyperlink');
    expect(tbl).toContain('a link');
    // The theme colours survive the rebuilt runs.
    expect(tbl).toContain('334155');
    expect(tbl).toContain('92400E');
    // No literal markdown markers leak into the text.
    expect(tbl).not.toContain('**');
    expect(tbl).not.toContain('`code`');
    const strict = await validateDocxBuffer(parser.rezip());
    expect(strict.ok).toBe(true);
    expect(strict.errors).toEqual([]);
  });

  it('leaves a plain alert byte-identical in structure (single body run)', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const result = await composeDocument(parser, '> [!NOTE]\n> **Title**\n> plain body\n');
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const tbl = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect((tbl.match(/<w:r>/g) ?? []).length).toBe(2);
  });

  it('joins several body paragraphs and resolves citations in the alert', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const md = [
      '> [!TIP]',
      '> First **part**.',
      '>',
      '> Second part cites [@smith].',
      '',
      '[@smith]: Smith, J. A Book. 2020.',
      '',
    ].join('\n');
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);
    const xml = documentXml(parser);
    const tbl = xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
    expect(tbl).toContain('Second part cites [1].');
  });

  it('still refuses images, raw HTML, inline math and note markers', () => {
    expect(() => compileMarkdown('> [!NOTE]\n> see ![a](x.png)\n')).toThrow(/images are out of scope/);
    expect(() => compileMarkdown('> [!NOTE]\n> a <b>x</b> b\n')).toThrow(/raw HTML/);
    expect(() => compileMarkdown('> [!NOTE]\n> solve \\(x+1\\) now\n')).toThrow(/inline math inside alerts/);
    expect(() => compileMarkdown('> [!NOTE]\n> note[^a]\n\n[^a]: body\n')).toThrow(/footnote markers inside alerts/);
  });

  it('refuses a link with a forbidden scheme inside an alert', () => {
    expect(() => compileMarkdown('> [!NOTE]\n> [x](javascript:alert(1))\n')).toThrow(/not allowed/);
  });

  it('lets a callout op carry a client id and refuses a duplicate id', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const engine = new OperationEngine(await parser.getAST(), parser);
    const first = await engine.executeBatch({
      operations: [{ action: 'callout', id: 'my_note', anchor: 'end', position: 'after', text: 'one', theme: 'indigo' }],
      sequential: true,
    });
    expect(first.success).toBe(true);
    const ast = await parser.getAST();
    expect(ast.sections.some((sec) => sec.nodes.some((n) => n.id === 'my_note'))).toBe(true);
    const second = await engine.executeBatch({
      operations: [{ action: 'callout', id: 'my_note', anchor: 'end', position: 'after', text: 'two', theme: 'indigo' }],
      sequential: true,
    });
    expect(second.success).toBe(false);
    expect(JSON.stringify(second)).toContain('collides');
  });
});
