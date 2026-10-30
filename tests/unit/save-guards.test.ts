/**
 * Save-guard proofs: duplicate native paraIds, validator authority over element names,
 * insert-into-footer session-id stripping, thrown-error envelopes, and
 * pending-normalization previews.
 *
 * Each describe maps to one engine-integrity fix; every markup assertion
 * reads live behavior (never assumptions). Async XSD oracles are used only
 * where the production path already requires them.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  compileMarkdown,
  createBlankDocxBuffer,
  executeInsert,
  getEngineRemedy,
} from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

function zipWithDocumentXml(xml: string): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
  );
  zip.addFile('word/document.xml', Buffer.from(xml, 'utf8'));
  return zip.toBuffer();
}

function dupParaIdBuffer(): Buffer {
  return zipWithDocumentXml(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}"><w:body>` +
      `<w:p w14:paraId="7B904439"><w:r><w:t>First</w:t></w:r></w:p>` +
      `<w:p w14:paraId="7B904439"><w:r><w:t>Second</w:t></w:r></w:p>` +
      `</w:body></w:document>`
  );
}

function bogusDocumentBuffer(): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const xml = zip.readAsText('word/document.xml');
  const at = xml.lastIndexOf('<w:sectPr');
  const bogus = '<w:p><w:r><w:bogus/></w:r></w:p>';
  zip.updateFile(
    'word/document.xml',
    Buffer.from(at >= 0 ? xml.slice(0, at) + bogus + xml.slice(at) : xml.replace('</w:body>', bogus + '</w:body>'), 'utf8')
  );
  return zip.toBuffer();
}

describe('duplicate native paraIds fail closed at save', () => {
  it('marks the repeat with a _dup session id instead of colliding', () => {
    const parser = new DocxParser(dupParaIdBuffer());
    const summary = parser.getStructuralSummary();
    expect(summary.nodes[0].id).toBe('7B904439');
    expect(summary.nodes[1].id).toBe('7B904439_dup');
  });

  it('names the paraId and xPath in the live scan', () => {
    const parser = new DocxParser(dupParaIdBuffer());
    const dups = parser.getDuplicateParaIds();
    expect(dups).toHaveLength(1);
    expect(dups[0].paraId).toBe('7B904439');
    expect(dups[0].count).toBe(2);
    expect(dups[0].xPaths).toEqual(['/w:document[1]/w:body[1]/w:p[2]']);
  });

  it('refuses bare rezip with a coded error naming the fix path', () => {
    const parser = new DocxParser(dupParaIdBuffer());
    try {
      parser.rezip();
      expect.unreachable();
    } catch (err: unknown) {
      const e = err as { code?: unknown; remedy?: unknown; message?: string };
      expect(e.code).toBe('E_VALIDATION');
      expect(typeof e.remedy === 'string' && e.remedy.length > 0).toBe(true);
      expect(String(e.message)).toContain('7B904439');
      expect(String(e.message)).toContain('/w:document[1]/w:body[1]/w:p[2]');
      expect(String(e.remedy)).toContain('7B904439_dup');
    }
  });

  it('reports ok:false (not a throw) through validateOOXML', async () => {
    const parser = new DocxParser(dupParaIdBuffer());
    const res = await parser.validateOOXML();
    expect(res.ok).toBe(false);
    const first = res.errors[0] as unknown as Record<string, unknown>;
    expect(first['errorType']).toBe('DuplicateParaId');
    expect(first['id']).toBe('7B904439');
    expect(first['xPath']).toBe('/w:document[1]/w:body[1]/w:p[2]');
    expect(first['xpath']).toBe('/w:document[1]/w:body[1]/w:p[2]');
  });

  it('leaves unique-id documents untouched (control)', () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    expect(parser.getDuplicateParaIds()).toEqual([]);
    expect(() => parser.rezip()).not.toThrow();
  });
});

describe('schema-invalid XML is refused by the validator, not by rezip', () => {
  it('bare rezip does not judge element names', () => {
    const parser = new DocxParser(bogusDocumentBuffer());
    const out = parser.rezip();
    expect(Buffer.isBuffer(out)).toBe(true);
    expect(new AdmZip(out).readAsText('word/document.xml')).toContain('w:bogus');
  });

  it('validateOOXML reports ok:false with the validator wording and an xPath', async () => {
    const parser = new DocxParser(bogusDocumentBuffer());
    const res = await parser.validateOOXML();
    expect(res.ok).toBe(false);
    const first = res.errors[0] as unknown as Record<string, unknown>;
    expect(String(first['description'])).toContain('bogus');
    expect(String(first['description'])).toContain('invalid child element');
    expect(first['xPath']).toMatch(/^\/w:document\[1\]\/w:body\[1\]\/w:p\[\d+\]\/w:r\[1\]$/);
  });

  it('assertValidDocx refuses the saved bytes with E_VALIDATION', async () => {
    const out = new DocxParser(bogusDocumentBuffer()).rezip();
    const { assertValidDocx } = await import('../../packages/engine/src/ooxml/validator-gate');
    await expect(assertValidDocx(out)).rejects.toMatchObject({ code: 'E_VALIDATION' });
  });

  it('offers an explicit opt-out that still succeeds', () => {
    const parser = new DocxParser(bogusDocumentBuffer());
    const out = parser.rezip({ validate: false });
    expect(Buffer.isBuffer(out)).toBe(true);
    expect(new AdmZip(out).readAsText('word/document.xml')).toContain('w:bogus');
  });

  it('saved bytes fail the async schema gate (layers agree)', async () => {
    const parser = new DocxParser(bogusDocumentBuffer());
    const out = parser.rezip({ validate: false });
    expect((await validateDocxBuffer(out)).ok).toBe(false);
  });
});

describe('insert paragraph into footer saves VALID without session ids', () => {
  it('strips the session id in the shared header/footer path', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const sections = (ast as unknown as { sections: Array<{ nodes: Array<{ id: string }> }> }).sections;
    const anchor = sections[0].nodes[0].id;
    await executeInsert(
      {
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Footer line',
        target: 'footer',
      } as unknown as Parameters<typeof executeInsert>[0],
      ast as unknown as Parameters<typeof executeInsert>[1],
      parser
    );
    const buf = parser.rezip();
    const footers = new AdmZip(buf)
      .getEntries()
      .filter((e) => /footer\d*\.xml$/.test(e.entryName));
    expect(footers.length).toBeGreaterThan(0);
    for (const f of footers) {
      expect(f.getData().toString('utf8')).not.toContain('data-astra-id');
    }
    expect((await validateDocxBuffer(buf)).ok).toBe(true);
  });
});

describe('thrown errors carry a single code/remedy envelope', () => {
  it('batch failures surface a coded critical entry (no message-parsing needed)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [{ action: 'replace', target: 'NOPE_MISSING', text: 'x' } as never],
      sequential: true,
    });
    expect(res.success).toBe(false);
    const critical = res.errors.find((e) => e.operationIndex === -1);
    expect(critical).toBeTruthy();
    // The internal rethrow keeps the machine code + remedy in text form so
    // every downstream mapper (CLI inferBatchCode, MCP envelopes) classifies
    // without substring-guessing uncoded failures.
    expect(String(critical!.error)).toContain('[E_VALIDATION] Batch operation 0 failed');
    expect(String(critical!.error)).toContain('Remedy:');
  });

  it('compose compile failures carry .code and .remedy', () => {
    try {
      compileMarkdown('   ');
      expect.unreachable();
    } catch (err: unknown) {
      const e = err as { code?: unknown; remedy?: unknown };
      expect(e.code).toBe('E_VALIDATION');
      expect(typeof e.remedy === 'string' && e.remedy.length > 0).toBe(true);
    }
  });

  it('schema-assertion throws carry .code and .remedy', async () => {
    const bogus = new DocxParser(bogusDocumentBuffer());
    const out = bogus.rezip({ validate: false });
    const { assertValidDocx } = await import('../../packages/engine/src/ooxml/validator-gate');
    try {
      await assertValidDocx(out);
      expect.unreachable();
    } catch (err: unknown) {
      const e = err as { code?: unknown; remedy?: unknown };
      expect(e.code).toBe('E_VALIDATION');
      expect(typeof e.remedy === 'string' && e.remedy.length > 0).toBe(true);
    }
  });

  it('the E_VALIDATION default remedy stays node-agnostic', () => {
    const remedy = getEngineRemedy('E_VALIDATION');
    expect(remedy).not.toContain('re-inspect ids');
    expect(remedy).toContain('re-export');
  });
});

describe('pending normalizations surface before save', () => {
  it('previews rich entries without mutating', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    await parser.getAST();
    // Misorder a live pPr directly: jc ahead of pStyle violates schema order.
    const live = parser.getDocumentDom();
    const pPr = live.getElementsByTagName('w:pPr')[0] as unknown as Element;
    const jc = live.createElementNS(WORD_NS, 'w:jc');
    jc.setAttribute('w:val', 'center');
    pPr.insertBefore(jc, pPr.firstChild);
    const first = parser.getPendingNormalizations();
    expect(first.length).toBeGreaterThan(0);
    const line = first.find((l) => l.startsWith('Reordered w:pPr into schema order'));
    expect(line).toBeTruthy();
    expect(line).toContain(' at /w:document[1]/w:body[1]/w:p[1]/w:pPr[1]');
    expect(line).toContain('(node ');
    expect(line).toContain('[w:jc,w:pStyle] \u2192 [w:pStyle,w:jc]');
    // Preview twice: identical (dry run heals nothing), then the real save
    // heals and the preview goes empty.
    expect(parser.getPendingNormalizations()).toEqual(first);
    parser.rezip();
    expect(parser.getPendingNormalizations()).toEqual([]);
  });
});
