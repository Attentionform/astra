import { describe, it, expect } from 'vitest';
import { DocxParser } from '../src/parser';
import { compileMarkdown, composeDocument } from '../src/compose/markdown';
import { validateBuffer } from '@attentionform/ooxml-validator';
import { OperationEngine } from '../src/operation-engine';
import type { InsertOperation } from '../src/operations/insert';

describe('Table Cell Rich Text and Hyperlinks (AS-03)', () => {
  it('preserves escaped Markdown delimiters as literal table text after reopening', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({ operations: [{ action: 'insert', type: 'table', anchor: 'default', position: 'after', headers: ['Literal'], rows: [[String.raw`\*\*urgent\*\*`]] }] });
    expect(res.success).toBe(true);
    const reopened = new DocxParser(parser.rezip());
    const table = reopened.getDocumentDom().getElementsByTagName('w:tbl')[0];
    expect(table.textContent).toContain('**urgent**');
  });

  it('compiles markdown table with bold, italic, code, strike, and hyperlinks without validation errors', () => {
    const md = [
      '| **Feature** | *Status* | `Code` | ~~Legacy~~ | [Link](https://example.com) |',
      '| --- | --- | --- | --- | --- |',
      '| Normal | **Bold text** | *Italic text* | ~~Strikethrough~~ | [ASTra Docs](https://astra.dev) |',
      '| Mixed | **`Bold Code`** | *`Italic Code`* | Plain | `inline_code()` |',
    ].join('\n');

    const ops = compileMarkdown(md);
    expect(ops.length).toBeGreaterThan(0);
    const tableOp = ops.find((op) => op.action === 'insert' && (op as { type?: string }).type === 'table') as any;
    expect(tableOp).toBeDefined();
    expect(tableOp.headers).toEqual([
      '**Feature**',
      '*Status*',
      '`Code`',
      '~~Legacy~~',
      '[Link](https://example.com)',
    ]);
    expect(tableOp.rows[0]).toEqual([
      'Normal',
      '**Bold text**',
      '*Italic text*',
      '~~Strikethrough~~',
      '[ASTra Docs](https://astra.dev)',
    ]);
  });

  it('composes markdown table with rich text and hyperlinks into valid OOXML with relationships', async () => {
    const parser = await DocxParser.createNewDocument('blank');
    const md = [
      '# Document Title',
      '',
      '| **Metric** | *Target* | `Reference` | Notes |',
      '| --- | --- | --- | --- |',
      '| Speed | **100%** | `v1.2.0` | See [Spec](https://specs.example.com) |',
      '| Legacy | ~~Deprecated~~ | `old_api` | Replaced by *new* system |',
    ].join('\n');

    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);
    expect(result.errors).toHaveLength(0);

    const docDom = parser.getDocumentDom();
    const tables = docDom.getElementsByTagName('w:tbl');
    expect(tables.length).toBe(1);
    const tbl = tables[0];

    // Check for bold elements
    const bEls = tbl.getElementsByTagName('w:b');
    expect(bEls.length).toBeGreaterThan(0);

    // Check for italic elements
    const iEls = tbl.getElementsByTagName('w:i');
    expect(iEls.length).toBeGreaterThan(0);

    // Check for strike elements
    const strikeEls = tbl.getElementsByTagName('w:strike');
    expect(strikeEls.length).toBeGreaterThan(0);

    // Check for code font elements
    const rFonts = Array.from(tbl.getElementsByTagName('w:rFonts'));
    const codeFonts = rFonts.filter(
      (rf) => rf.getAttribute('w:ascii') === 'Courier New' && rf.getAttribute('w:hAnsi') === 'Courier New'
    );
    expect(codeFonts.length).toBeGreaterThan(0);

    // Check for hyperlink element
    const hyperlinks = tbl.getElementsByTagName('w:hyperlink');
    expect(hyperlinks.length).toBe(1);
    const relId = hyperlinks[0].getAttribute('r:id');
    expect(relId).toBeTruthy();

    // Verify relationship in word/_rels/document.xml.rels
    const zip = parser.getZipAdapter();
    const relsXml = zip.readText('word/_rels/document.xml.rels');
    expect(relsXml).toContain(relId as string);
    expect(relsXml).toContain('https://specs.example.com');
    expect(relsXml).toContain('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');

    // Verify OOXML validity
    const buf = await parser.rezip();
    const validation = await validateBuffer(buf);
    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });

  it('renders rich text and hyperlinks via direct batch insert table operation', async () => {
    const parser = await DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const op: InsertOperation = {
      action: 'insert',
      type: 'table',
      anchor: 'default',
      position: 'after',
      headers: ['**Header 1**', '[Website](https://mywebsite.com)'],
      rows: [
        ['`code_cell`', 'Plain text with *italics* and ~~strike~~'],
      ],
    };

    const res = await engine.executeBatch({ operations: [op], sequential: true });
    expect(res.success).toBe(true);

    const docDom = parser.getDocumentDom();
    const tables = docDom.getElementsByTagName('w:tbl');
    expect(tables.length).toBe(1);

    const tbl = tables[0];
    const hyperlinks = tbl.getElementsByTagName('w:hyperlink');
    expect(hyperlinks.length).toBe(1);
    expect(hyperlinks[0].textContent).toBe('Website');

    const zip = parser.getZipAdapter();
    const relsXml = zip.readText('word/_rels/document.xml.rels');
    expect(relsXml).toContain('https://mywebsite.com');

    const buf = await parser.rezip();
    const validation = await validateBuffer(buf);
    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });

  it('rejects images inside table cells', () => {
    const md = [
      '| Col 1 | Col 2 |',
      '| --- | --- |',
      '| ![Image](test.png) | Data |',
    ].join('\n');

    expect(() => compileMarkdown(md)).toThrow(/images are out of scope in a table/);
  });

  it('rejects raw HTML inside table cells', () => {
    const md = [
      '| Col 1 | Col 2 |',
      '| --- | --- |',
      '| <div>Hello</div> | Data |',
    ].join('\n');

    expect(() => compileMarkdown(md)).toThrow(/raw HTML is out of scope in a table/);
  });

  it('rejects inline math inside table cells', () => {
    const md = [
      '| Col 1 | Col 2 |',
      '| --- | --- |',
      '| $E = mc^2$ | Data |',
    ].join('\n');

    expect(() => compileMarkdown(md, { dollarMath: true })).toThrow(/inline math inside tables is out of scope/);
  });

  it('rejects disallowed link schemes inside table cells', () => {
    const md = [
      '| Col 1 | Col 2 |',
      '| --- | --- |',
      '| [Link](javascript:alert(1)) | Data |',
    ].join('\n');

    expect(() => compileMarkdown(md)).toThrow(/link scheme 'javascript:' is not allowed/);
  });

  it('rejects nested links inside table cells', () => {
    const md = [
      '| Col 1 | Col 2 |',
      '| --- | --- |',
      '| [[Inner](https://a.com)](https://b.com) | Data |',
    ].join('\n');

    expect(() => compileMarkdown(md)).toThrow(/nested links are out of scope/);
  });
});
