import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  InsertOperation,
  OperationEngine,
  SectionOperation,
  StyleOperation,
  convertAstToHtml,
  convertAstToMarkdown,
} from '@attentionform/astra-engine';

import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';

function createMinimalDocx(): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  );
  zip.addFile(
    'word/document.xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p data-astra-id="11111111" w14:paraId="11111111">
      <w:r><w:t>Section 1 Content</w:t></w:r>
    </w:p>
    <w:p data-astra-id="22222222" w14:paraId="22222222">
      <w:r><w:t>Section 2 Content</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Page Separation & Divider Engine Suite', () => {
  it('inserts horizontal section divider line with custom double lineStyle, color, and spacing', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const divOp: InsertOperation = {
      action: 'insert',
      type: 'divider',
      anchor: '11111111',
      position: 'after',
      lineStyle: 'double',
      color: '4F46E5',
      thickness: 18,
      spacingBefore: 180,
      spacingAfter: 180
    };

    const result = await engine.executeBatch({ operations: [divOp] });
    expect(result.success).toBe(true);

    const dividerNode = ast.sections[0].nodes[1];
    expect(dividerNode).toBeDefined();

    const xmlStr = new XMLSerializer().serializeToString(dividerNode.xml);
    expect(xmlStr).toContain('w:pBdr');
    expect(xmlStr).toContain('w:bottom w:val="double" w:sz="18" w:space="1" w:color="4F46E5"');
    expect(xmlStr).toContain('w:spacing w:before="180" w:after="180"');
  });

  it('enables vertical column separator in multi-column sections (w:sep="1")', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const secOp: SectionOperation = {
      action: 'section',
      operation: 'setColumns',
      columns: 2,
      spacing: 720,
      separator: true
    };

    const result = await engine.executeBatch({ operations: [secOp] });
    expect(result.success).toBe(true);

    const doc = ast.sections[0].nodes[0].xml.ownerDocument;
    const docXml = new XMLSerializer().serializeToString(doc!);
    expect(docXml).toContain('w:cols');
    expect(docXml).toContain('w:num="2"');
    expect(docXml).toContain('w:sep="1"');
  });

  it('applies paragraph borders (borderBottom, borderLeft, borderBox)', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const styleOp: StyleOperation = {
      action: 'style',
      target: '11111111',
      properties: {
        borderBottom: { style: 'single', sz: 12, color: '6366F1' },
        borderLeft: { style: 'single', sz: 32, color: '4F46E5' }
      }
    };

    const result = await engine.executeBatch({ operations: [styleOp] });
    expect(result.success).toBe(true);

    const node = ast.sections[0].nodes.find(n => n.id === '11111111');
    const xmlStr = new XMLSerializer().serializeToString(node!.xml);
    expect(xmlStr).toContain('w:pBdr');
    expect(xmlStr).toContain('w:bottom w:val="single" w:sz="12" w:space="1" w:color="6366F1"');
    expect(xmlStr).toContain('w:left w:val="single" w:sz="32" w:space="1" w:color="4F46E5"');
  });
});

describe('Multi-line paragraph inserts split newlines into break runs', () => {
  async function insertParagraph(id: string, text: string) {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const op: InsertOperation = {
      action: 'insert', type: 'paragraph', anchor: '11111111', position: 'after', text,
    };
    (op as any).id = id;
    const result = await engine.executeBatch({ operations: [op] });
    expect(result.success).toBe(true);
    const node = ast.sections[0].nodes.find(n => n.id === id);
    expect(node).toBeDefined();
    return { ast, node: node! };
  }

  it('splits embedded newlines into one run per line with w:br between lines', async () => {
    const { node } = await insertParagraph('multi-1', 'line1\nline2\nline3');
    const xmlStr = new XMLSerializer().serializeToString(node.xml);
    expect((xmlStr.match(/<w:r[\s>]/g) || []).length).toBe(5);
    expect((xmlStr.match(/<w:br[\s/>]/g) || []).length).toBe(2);
    expect(xmlStr).toContain('<w:t>line1</w:t>');
    expect(xmlStr).toContain('<w:t>line3</w:t>');
    expect(/<w:t>[^<]*(?:\r|\n)/.test(xmlStr)).toBe(false);
  });

  it('handles CRLF and lone CR separators and leaves single-line text as one run', async () => {
    const { node } = await insertParagraph('multi-2', 'a\r\nb\rc');
    const xmlStr = new XMLSerializer().serializeToString(node.xml);
    expect((xmlStr.match(/<w:br[\s/>]/g) || []).length).toBe(2);
    expect(xmlStr).toContain('<w:t>a</w:t>');
    const single = await insertParagraph('single-1', 'plain');
    const singleXml = new XMLSerializer().serializeToString(single.node.xml);
    expect((singleXml.match(/<w:r[\s>]/g) || []).length).toBe(1);
    expect(singleXml).toContain('<w:t>plain</w:t>');
  });

  it('preserves line breaks through Markdown and HTML export', async () => {
    const { ast } = await insertParagraph('multi-3', 'alpha\nbeta');
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('alpha\nbeta');
    const html = convertAstToHtml(ast, 'multiline');
    expect(html).toContain('alpha<br/>beta');
  });
});
