import { describe, it, expect, beforeEach } from 'vitest';
import { DocxParser } from '@attentionform/astra-engine';
import AdmZip from 'adm-zip';

describe('Part 2: DOCX Parsing & DOM Extraction', () => {
  let mockDocxBuffer: Buffer;

  beforeEach(() => {
    // Generate an in-memory zip file containing a mock word/document.xml
    const zip = new AdmZip();
    const mockXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="3A0108A4">
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r><w:t>Hello World</w:t></w:r>
    </w:p>
    <w:p w14:paraId="22222222">
      <w:r><w:t>This is a normal paragraph.</w:t></w:r>
    </w:p>
    <w:p>
      <w:ins><w:r><w:t>Tracked insertion</w:t></w:r></w:ins>
    </w:p>
  </w:body>
</w:document>`;
    zip.addFile('word/document.xml', Buffer.from(mockXml, 'utf8'));
    // Add dummy content types and rels to ensure zip isn't totally empty
    zip.addFile('[Content_Types].xml', Buffer.from('<Types></Types>'));
    mockDocxBuffer = zip.toBuffer();
  });

  it('should parse and extract Structural Summary correctly', () => {
    const parser = new DocxParser(mockDocxBuffer);
    const summary = parser.getStructuralSummary();
    
    expect(summary.hasTrackedChanges).toBe(true);
    expect(summary.nodes.length).toBe(3);
    
    // First node should use the native w14:paraId
    expect(summary.nodes[0].id).toBe('3A0108A4');
    expect(summary.nodes[0].isHeading).toBe(true);
    expect(summary.nodes[0].preview).toBe('Hello World');

    // Second node reuses its native w14:paraId
    expect(summary.nodes[1].id).toBe('22222222');
    expect(summary.nodes[1].isHeading).toBe(false);
  });

  it('should read a specific node', () => {
    const parser = new DocxParser(mockDocxBuffer);
    const nodeXml = parser.readNode('22222222');
    expect(nodeXml).toContain('This is a normal paragraph.');
  });

  it('should reject XML injection and path traversal attempts', () => {
    const parser = new DocxParser(mockDocxBuffer);
    
    // Path traversal test on ID
    expect(() => parser.readNode('../../word/styles.xml')).toThrow(/Invalid Node ID format/);

    // XML Injection test on updateNode
    // Actually xmldom will parse <script> as valid XML tags if it's well formed.
    // But unbalanced tags or actual strict XML injection will fail.
    const unbalancedXml = `<w:p><w:t>Inject</w:t><</w:p>`;
    expect(() => parser.updateNode('22222222', unbalancedXml)).toThrow();
  });

  it('should achieve zero data-loss on semantic rezip', () => {
    const parser = new DocxParser(mockDocxBuffer);
    const rezipppedBuffer = parser.rezip();
    
    // Unzip the new buffer and compare document.xml
    const newZip = new AdmZip(rezipppedBuffer);
    const newXml = newZip.getEntry('word/document.xml')?.getData().toString('utf8');
    
    // Ensure our temporary data-astra-id attribute didn't leak into the final output
    expect(newXml).not.toContain('data-astra-id');
    // Ensure original native attributes like w14:paraId are preserved perfectly
    expect(newXml).toContain('w14:paraId="3A0108A4"');
  });
});
