import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  FieldOperation,
  OperationEngine,
  convertDocxToPdf,
  hasExportPdfBackend,
} from '@attentionform/astra-engine';

import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';
import { execSync } from 'node:child_process';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { withLoSlot } from '../helpers/lo-slot';
import { useReproducibleOutput } from '../helpers/reproducible';

// Saved buffers are compared byte for byte: pin the zip and revision timestamps.
useReproducibleOutput();

// Slot + budget note: the live LO render below holds the suite-wide LO slot
// (see tests/helpers/lo-slot.ts) and carries 300s for queue + own run.

function createHeadingDocx(): Buffer {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p data-astra-id="p_title"><w:r><w:t>Annual Report</w:t></w:r></w:p>
    <w:p data-astra-id="p_h1">
      <w:pPr><w:pStyle w:val="Heading 1"/></w:pPr>
      <w:r><w:t>Executive Summary</w:t></w:r>
    </w:p>
    <w:p data-astra-id="p_h2">
      <w:pPr><w:pStyle w:val="Heading 2"/></w:pPr>
      <w:r><w:t>Financial Highlights</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Field Operations Suite (TOC & Field Codes)', () => {
  it('generates dynamic Table of Contents (TOC) with dot leaders and field codes', async () => {
    const buffer = createHeadingDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const tocOp: FieldOperation = {
      action: 'field',
      operation: 'insertTOC',
      anchor: 'p_title',
      position: 'after',
      title: 'Document Outline & Contents'
    };

    const result = await engine.executeBatch({ operations: [tocOp] });
    expect(result.success).toBe(true);

    const tocNode = ast.sections[0].nodes.find(n => n.id.startsWith('toc_'));
    expect(tocNode).toBeDefined();

    const xmlStr = new XMLSerializer().serializeToString(tocNode!.xml);
    expect(xmlStr).toContain('TOC \\o &quot;1-3&quot; \\h \\z \\u');
    expect(xmlStr).toContain('Document Outline &amp; Contents');
  });

  it('inserts page number field code', async () => {
    const buffer = createHeadingDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const pageOp: FieldOperation = {
      action: 'field',
      operation: 'insertPageNumber',
      anchor: 'p_title',
      position: 'before'
    };

    const result = await engine.executeBatch({ operations: [pageOp] });
    expect(result.success).toBe(true);

    const pageNode = ast.sections[0].nodes.find(n => n.id.startsWith('field_'));
    const xmlStr = new XMLSerializer().serializeToString(pageNode!.xml);
    expect(xmlStr).toContain('w:instr="PAGE"');
  });
});

// Header/footer-targeted field + insert ops. A page-number field in the body
// renders once where it sits; in the footer part it repeats on every page.
function createMultipageDocx(pages: number): Buffer {
  const paras: string[] = [
    '<w:p data-astra-id="p_title"><w:r><w:t>Annual Report</w:t></w:r></w:p>'
  ];
  for (let i = 1; i <= pages; i++) {
    paras.push(`<w:p data-astra-id="p_${i}"><w:r><w:t>Body content page ${i}</w:t></w:r></w:p>`);
    if (i < pages) {
      paras.push(`<w:p data-astra-id="pb_${i}"><w:r><w:br w:type="page"/></w:r></w:p>`);
    }
  }
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
    '_rels/.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`)
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  );
  zip.addFile(
    'word/document.xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${paras.join('')}
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

function readPart(buffer: Buffer, re: RegExp): string | null {
  const zip = new AdmZip(buffer);
  const entry = zip.getEntries().find(e => re.test(e.entryName));
  return entry ? zip.readAsText(entry) : null;
}

function hasBackend(): boolean {
  try {
    if (!hasExportPdfBackend()) return false;
    execSync('which pdftotext', { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
}

describe('Header/footer-targeted field and insert ops', () => {
  it('routes a page-number field to the footer (not the body)', async () => {
    const parser = new DocxParser(createMultipageDocx(2));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const bodyCount = ast.sections[0].nodes.length;

    const result = await engine.executeBatch({
      operations: [{ action: 'field', operation: 'insertPageNumber', target: 'footer' } as FieldOperation]
    });
    expect(result.success).toBe(true);

    // No new body node: the failure mode this covers was body placement.
    expect(ast.sections[0].nodes.length).toBe(bodyCount);
    expect(ast.sections[0].nodes.some(n => n.id.startsWith('field_'))).toBe(false);

    const saved = parser.rezip();
    const footer = readPart(saved, /footer\d+\.xml$/);
    expect(footer).not.toBeNull();
    expect(footer!).toContain('PAGE');
    expect(footer!).toContain('NUMPAGES');
    const docXml = readPart(saved, /word\/document\.xml$/)!;
    expect(docXml).not.toContain('PAGE');
    expect(docXml).toContain('w:footerReference');
  });

  it('accepts the footer target nested under params', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'field',
        operation: 'insertPageNumber',
        params: { target: 'footer' }
      } as unknown as FieldOperation]
    });
    expect(result.success).toBe(true);
    const footer = readPart(parser.rezip(), /footer\d+\.xml$/);
    expect(footer).not.toBeNull();
    expect(footer!).toContain('PAGE');
  });

  it('routes a word-count field to the header', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'field',
        operation: 'insertWordCount',
        params: { target: 'header' }
      } as unknown as FieldOperation]
    });
    expect(result.success).toBe(true);
    const saved = parser.rezip();
    const header = readPart(saved, /header\d+\.xml$/);
    expect(header).not.toBeNull();
    expect(header!).toContain('NUMWORDS');
    expect(readPart(saved, /word\/document\.xml$/)!).not.toContain('NUMWORDS');
  });

  it('appends a second footer field into the same part without wiping the first', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const first = await engine.executeBatch({
      operations: [{ action: 'field', operation: 'insertPageNumber', target: 'footer' } as FieldOperation]
    });
    expect(first.success).toBe(true);
    const second = await engine.executeBatch({
      operations: [{
        action: 'insert', type: 'paragraph', anchor: 'p_title', position: 'after',
        text: 'Confidential', target: 'footer'
      } as any]
    });
    expect(second.success).toBe(true);

    const footer = readPart(parser.rezip(), /footer\d+\.xml$/)!;
    expect(footer).toContain('PAGE');
    expect(footer).toContain('Confidential');
    expect(footer.indexOf('PAGE')).toBeLessThan(footer.indexOf('Confidential'));
  });

  it('fails closed on an unknown field target with remedy and no mutation', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const before = Buffer.from(parser.rezip()).toString('base64');

    const result = await engine.executeBatch({
      operations: [{ action: 'field', operation: 'insertPageNumber', target: 'sidebar' } as unknown as FieldOperation]
    });
    expect(result.success).toBe(false);
    expect(result.errors[0]?.error).toMatch(/E_VALIDATION/);
    expect(result.errors[0]?.error).toMatch(/Remedy/);
    expect(Buffer.from(parser.rezip()).toString('base64')).toBe(before);
  });

  it('fails closed on an unknown pageType', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'field', operation: 'insertPageNumber', target: 'footer', pageType: 'middle'
      } as unknown as FieldOperation]
    });
    expect(result.success).toBe(false);
    expect(result.errors[0]?.error).toMatch(/E_VALIDATION.*pageType/);
  });

  it('fails closed when a TOC targets the footer', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'field', operation: 'insertTOC', anchor: 'p_title', target: 'footer'
      } as unknown as FieldOperation]
    });
    expect(result.success).toBe(false);
    expect(result.errors[0]?.error).toMatch(/E_VALIDATION/);
  });

  it('inserts paragraph text into the footer and leaves the body untouched', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'insert', type: 'paragraph', anchor: 'p_title', position: 'after',
        text: 'Draft — do not distribute', target: 'footer'
      } as any]
    });
    expect(result.success).toBe(true);
    const saved = parser.rezip();
    expect(readPart(saved, /footer\d+\.xml$/)!).toContain('Draft — do not distribute');
    expect(readPart(saved, /word\/document\.xml$/)!).not.toContain('Draft — do not distribute');
  });

  it('fails closed when a table targets the footer, with a working alternative', async () => {
    const parser = new DocxParser(createMultipageDocx(1));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'insert', type: 'table', anchor: 'p_title', position: 'after', target: 'footer'
      } as any]
    });
    expect(result.success).toBe(false);
    expect(result.errors[0]?.error).toMatch(/E_VALIDATION/);
    expect(result.errors[0]?.error).toMatch(/headerFooter/);
  });

  it('footer page-number part carries no session tracking attribute and passes validation', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{ action: 'field', operation: 'insertPageNumber', target: 'footer' } as FieldOperation]
    });
    expect(result.success).toBe(true);

    const saved = parser.rezip();
    const footer = readPart(saved, /footer\d+\.xml$/);
    expect(footer).not.toBeNull();
    expect(footer!).toContain('PAGE');
    expect(footer!).toContain('NUMPAGES');
    expect(footer!).not.toContain('data-astra-id');

    const validation = await parser.validateOOXML();
    expect(validation.errors).toEqual([]);
    expect(validation.ok).toBe(true);
  });

  it('header word-count part carries no session tracking attribute and passes validation', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [{
        action: 'field',
        operation: 'insertWordCount',
        params: { target: 'header' }
      } as unknown as FieldOperation]
    });
    expect(result.success).toBe(true);

    const saved = parser.rezip();
    const header = readPart(saved, /header\d+\.xml$/);
    expect(header).not.toBeNull();
    expect(header!).toContain('NUMWORDS');
    expect(header!).not.toContain('data-astra-id');

    const validation = await parser.validateOOXML();
    expect(validation.errors).toEqual([]);
    expect(validation.ok).toBe(true);
  });

  it('body page-number insert keeps its session id and still passes validation', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;

    const result = await engine.executeBatch({
      operations: [{
        action: 'field', operation: 'insertPageNumber', anchor, position: 'after'
      } as FieldOperation]
    });
    expect(result.success).toBe(true);

    const pageNode = ast.sections[0].nodes.find(n => n.id.startsWith('field_'));
    expect(pageNode).toBeDefined();
    expect(new XMLSerializer().serializeToString(pageNode!.xml)).toContain('data-astra-id');

    const validation = await parser.validateOOXML();
    expect(validation.errors).toEqual([]);
    expect(validation.ok).toBe(true);
  });

  it('renders the footer page number on every page of the exported PDF', async () => {
    if (!hasBackend()) {
      console.warn('SKIP live PDF proof: no LibreOffice/pdftotext backend in this environment');
      return;
    }
    const parser = new DocxParser(createMultipageDocx(3));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const result = await engine.executeBatch({
      operations: [{ action: 'field', operation: 'insertPageNumber', target: 'footer' } as FieldOperation]
    });
    expect(result.success).toBe(true);

    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'field-footer-proof-'));
    const pdfPath = path.join(dir, 'proof.pdf');
    await withLoSlot(() => convertDocxToPdf(parser.rezip(), pdfPath));
    const text = execSync(`pdftotext -layout "${pdfPath}" -`, { encoding: 'utf8' });
    const pages = text.split('\f').map(p => p.trim()).filter(Boolean);
    expect(pages.length).toBe(3);
    for (let i = 0; i < pages.length; i++) {
      expect(pages[i]).toContain(`Page ${i + 1} of 3`);
    }
    fs.rmSync(dir, { recursive: true, force: true });
  }, 300000);
});
