/**
 * Phase E2 — document core-properties (metadata) read/write parity.
 *
 * F-1 ingest docProps/core.xml -> ast.metadata (never throws when absent);
 * F-2 read/update/synthesize the part (+ Content_Types Override + pkg rel);
 * F-3 {action:'metadata',operation:'setProperties'} partial-update op;
 * F-4 execute_batch exposure + inspect metadata surfacing.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  OperationEngine,
  executeMetadata,
  validateOperationsStrict,
  parseCorePropertiesXml,
  mergeCorePropertiesXml,
  CorePropertiesManipulator,
} from '@attentionform/astra-engine';
import {
  handleExecuteBatch,
  handleInspectDocument,
} from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';

const KNOWN_CORE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
  `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
  `xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
  `<dc:title>E2 Fixture Title</dc:title>` +
  `<dc:subject>E2 Subject</dc:subject>` +
  `<dc:creator>E2 Author</dc:creator>` +
  `<cp:keywords>alpha, beta</cp:keywords>` +
  `<dc:description>E2 Description</dc:description>` +
  `<cp:lastModifiedBy>E2 Editor</cp:lastModifiedBy>` +
  `<cp:revision>7</cp:revision>` +
  `<cp:category>E2 Category</cp:category>` +
  `<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-02T03:04:05Z</dcterms:created>` +
  `<dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-06T07:08:09Z</dcterms:modified>` +
  `</cp:coreProperties>`;

/** Blank docx buffer with KNOWN_CORE_XML injected as docProps/core.xml. */
function bufferWithKnownCore(): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.addFile('docProps/core.xml', Buffer.from(KNOWN_CORE_XML, 'utf8'));
  return zip.toBuffer();
}

/** A new document carries core properties; take them out, with their relationship and content type. */
function withoutCoreProperties(buffer: Buffer): Buffer {
  const zip = new AdmZip(buffer);
  zip.deleteFile('docProps/core.xml');
  const rels = zip
    .readAsText('_rels/.rels')
    .replace(/<Relationship\b[^>]*core-properties[^>]*\/>/g, '');
  zip.updateFile('_rels/.rels', Buffer.from(rels, 'utf8'));
  const types = zip
    .readAsText('[Content_Types].xml')
    .replace(/<Override\b[^>]*\/docProps\/core\.xml[^>]*\/>/g, '');
  zip.updateFile('[Content_Types].xml', Buffer.from(types, 'utf8'));
  return zip.toBuffer();
}

describe('E2 F-1 ingest: docProps/core.xml -> ast.metadata', () => {
  it('reads known core.xml values with correct types', async () => {
    const parser = new DocxParser(bufferWithKnownCore());
    const ast = await parser.getAST();
    expect(ast.metadata.title).toBe('E2 Fixture Title');
    expect(ast.metadata.subject).toBe('E2 Subject');
    expect(ast.metadata.creator).toBe('E2 Author');
    expect(ast.metadata.keywords).toBe('alpha, beta');
    expect(ast.metadata.description).toBe('E2 Description');
    expect(ast.metadata.lastModifiedBy).toBe('E2 Editor');
    expect(ast.metadata.revision).toBe(7);
    expect(ast.metadata.category).toBe('E2 Category');
    expect(ast.metadata.created).toBeInstanceOf(Date);
    expect(ast.metadata.modified).toBeInstanceOf(Date);
    expect((ast.metadata.created as Date).toISOString()).toBe('2026-01-02T03:04:05.000Z');
    expect((ast.metadata.modified as Date).toISOString()).toBe('2026-05-06T07:08:09.000Z');
  });

  it('pure parse helper maps namespaces and skips absent fields', () => {
    const meta = parseCorePropertiesXml(KNOWN_CORE_XML);
    expect(meta.title).toBe('E2 Fixture Title');
    expect(meta.revision).toBe(7);
    const minimal = parseCorePropertiesXml(
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">` +
        `<dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">Solo</dc:creator></cp:coreProperties>`
    );
    expect(minimal.creator).toBe('Solo');
    expect(minimal.title).toBeUndefined();
    expect(minimal.revision).toBeUndefined();
  });

  it('missing core.xml never throws and leaves fields undefined', async () => {
    const parser = new DocxParser(withoutCoreProperties(createBlankDocxBuffer('blank')));
    expect(parser.getZipAdapter().hasFile('docProps/core.xml')).toBe(false);
    const ast = await parser.getAST();
    expect(ast.metadata.title).toBeUndefined();
    expect(ast.metadata.creator).toBeUndefined();
    expect(ast.metadata.lastModifiedBy).toBeUndefined();
    expect(ast.metadata.revision).toBeUndefined();
    expect(ast.metadata.created).toBeUndefined();
    expect(ast.metadata.modified).toBeUndefined();
    expect(ast.metadata.subject).toBeUndefined();
    expect(ast.metadata.keywords).toBeUndefined();
  });
});

describe('E2 F-2/F-3 setProperties round-trip save -> reload -> read', () => {
  it('round-trips title/creator/keywords through save and reload', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [
        {
          action: 'metadata',
          operation: 'setProperties',
          title: 'Round Trip Title',
          creator: 'Round Trip Author',
          keywords: 'round, trip',
        } as never,
      ],
      sequential: true,
    });
    expect(res.success).toBe(true);
    expect(res.appliedCount).toBe(1);
    expect(ast.metadata.title).toBe('Round Trip Title');

    const buf = parser.rezip(ast);
    const zip = new AdmZip(buf);
    expect(zip.getEntry('docProps/core.xml')).not.toBeNull();
    const types = zip.readAsText('[Content_Types].xml');
    expect(types).toContain('application/vnd.openxmlformats-package.core-properties+xml');
    expect(types).toContain('/docProps/core.xml');
    const rels = zip.readAsText('_rels/.rels');
    expect(rels).toContain('metadata/core-properties');
    expect(rels).toContain('docProps/core.xml');

    const reloaded = await new DocxParser(buf).getAST();
    expect(reloaded.metadata.title).toBe('Round Trip Title');
    expect(reloaded.metadata.creator).toBe('Round Trip Author');
    expect(reloaded.metadata.keywords).toBe('round, trip');
  });

  it('partial update preserves unprovided fields (AST + part + reload)', async () => {
    const parser = new DocxParser(bufferWithKnownCore());
    const ast = await parser.getAST();
    expect(ast.metadata.title).toBe('E2 Fixture Title');

    const out = (await executeMetadata(
      { action: 'metadata', operation: 'setProperties', subject: 'New Subject' } as never,
      ast,
      parser
    )) as { updated: string[] };
    expect(out.updated).toEqual(['subject']);

    expect(ast.metadata.subject).toBe('New Subject');
    expect(ast.metadata.title).toBe('E2 Fixture Title');
    expect(ast.metadata.creator).toBe('E2 Author');
    expect(ast.metadata.revision).toBe(7);

    // Unknown siblings + untouched fields survive in the serialized part.
    const partXml = parser.getZipAdapter().readText('docProps/core.xml');
    expect(partXml).toContain('E2 Fixture Title');
    expect(partXml).toContain('New Subject');

    const reloaded = await new DocxParser(parser.rezip(ast)).getAST();
    expect(reloaded.metadata.subject).toBe('New Subject');
    expect(reloaded.metadata.title).toBe('E2 Fixture Title');
    expect(reloaded.metadata.creator).toBe('E2 Author');
    expect(reloaded.metadata.keywords).toBe('alpha, beta');
  });

  it('empty setProperties and unknown metadata operations fail loudly', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    await expect(
      executeMetadata({ action: 'metadata', operation: 'setProperties' } as never, ast, parser)
    ).rejects.toThrow(/E_VALIDATION/);
    await expect(
      executeMetadata({ action: 'metadata', operation: 'nope' } as never, ast, parser)
    ).rejects.toThrow(/E_VALIDATION/);
    // Strict schema rejects unknown flat keys and unknown operations.
    expect(() =>
      validateOperationsStrict([{ action: 'metadata', operation: 'setProperties', bogus: 1 }])
    ).toThrow(/E_VALIDATION/);
    expect(() =>
      validateOperationsStrict([{ action: 'metadata', operation: 'delete' }])
    ).toThrow(/E_VALIDATION/);
  });
});

describe('E2 F-4 MCP: execute_batch + inspect surface metadata', () => {
  it('metadata op flows through the operation pipeline and inspect shows title/creator', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'e2-meta.docx');

    const batch = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [
        {
          action: 'metadata',
          operation: 'setProperties',
          title: 'MCP Title',
          creator: 'MCP Author',
          category: 'MCP Cat',
        },
      ],
      sequential: true,
    });
    expect(batch).toMatchObject({ ok: true, appliedCount: 1 });

    const inspect = await handleInspectDocument(store, {
      sessionId: handle.sessionId,
      fidelity: 'standard',
    });
    expect(inspect).toMatchObject({ ok: true });
    if (inspect.ok !== true) throw new Error('inspect failed');
    expect(inspect.metadata.title).toBe('MCP Title');
    expect(inspect.metadata.creator).toBe('MCP Author');
    expect(inspect.metadata.category).toBe('MCP Cat');
  });

  it('inspect on a doc without core.xml reports empty metadata (no throw)', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(withoutCoreProperties(createBlankDocxBuffer('blank')), 'e2-nometa.docx');
    const inspect = await handleInspectDocument(store, {
      sessionId: handle.sessionId,
      fidelity: 'standard',
    });
    expect(inspect).toMatchObject({ ok: true });
    if (inspect.ok !== true) throw new Error('inspect failed');
    expect(inspect.metadata).toEqual({});
  });
});

describe('E2 loop-back: reviewer items 1-9', () => {
  it('T1 non-sequential: metadata (rank 0) runs before insert despite input order', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = ast.sections[0]?.nodes[0]?.id;
    expect(anchor).toBeDefined();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: anchor as string, position: 'after', text: 'sort-probe' } as never,
        { action: 'metadata', operation: 'setProperties', title: 'Sort Title' } as never,
      ],
      sequential: false,
    });
    expect(res.success).toBe(true);
    const hist = engine.getHistory();
    expect(hist.length).toBe(2);
    expect(hist[0]?.action).toBe('metadata');
    expect(hist[1]?.action).toBe('insert');
    expect(ast.metadata.title).toBe('Sort Title');
  });

  it('T2 reconcile: in-memory title survives rezip onto a stale part', async () => {
    const parser = new DocxParser(bufferWithKnownCore());
    const ast = await parser.getAST();
    expect(ast.metadata.title).toBe('E2 Fixture Title');
    ast.metadata.title = 'Memory New Title';
    const buf2 = parser.rezip(ast);
    const partXml = new AdmZip(buf2).readAsText('docProps/core.xml');
    expect(partXml).toContain('Memory New Title');
    const reloaded = await new DocxParser(buf2).getAST();
    expect(reloaded.metadata.title).toBe('Memory New Title');
    expect(reloaded.metadata.creator).toBe('E2 Author');
  });

  it('T3 malformed core.xml is loud E_VALIDATION (parse/merge/getAST)', async () => {
    expect(() => parseCorePropertiesXml('<cp:coreProperties><dc:title>oops')).toThrow(/E_VALIDATION/);
    expect(() =>
      mergeCorePropertiesXml('<cp:coreProperties><dc:title>oops', { title: 'x' })
    ).toThrow(/E_VALIDATION/);
    const bad = new AdmZip(createBlankDocxBuffer('blank'));
    bad.addFile('docProps/core.xml', Buffer.from('<cp:coreProperties><dc:title>oops', 'utf8'));
    await expect(new DocxParser(bad.toBuffer()).getAST()).rejects.toThrow(/E_VALIDATION/);
  });

  it('T4 prefix-less parts resolve via NS-URI/localName fallback', async () => {
    const prefixless =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">` +
      `<title xmlns="http://purl.org/dc/elements/1.1/">NoPrefix Title</title>` +
      `<creator xmlns="http://purl.org/dc/elements/1.1/">NoPrefix Author</creator>` +
      `</coreProperties>`;
    const meta = parseCorePropertiesXml(prefixless);
    expect(meta.title).toBe('NoPrefix Title');
    expect(meta.creator).toBe('NoPrefix Author');
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.addFile('docProps/core.xml', Buffer.from(prefixless, 'utf8'));
    const ast = await new DocxParser(zip.toBuffer()).getAST();
    expect(ast.metadata.title).toBe('NoPrefix Title');
    expect(ast.metadata.creator).toBe('NoPrefix Author');
  });

  it('T5 lastPrinted ranks revision<lastPrinted<created; category/contentStatus tail; unknown preserved', async () => {
    const withExtras = KNOWN_CORE_XML.replace(
      '<dcterms:created',
      '<cp:lastPrinted>3</cp:lastPrinted><dcterms:created'
    ).replace(
      '</cp:coreProperties>',
      '<cp:contentStatus>Draft</cp:contentStatus><my:custom xmlns:my="http://example.com/e2">keepme</my:custom></cp:coreProperties>'
    );
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.addFile('docProps/core.xml', Buffer.from(withExtras, 'utf8'));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    await executeMetadata(
      { action: 'metadata', operation: 'setProperties', title: 'Ordered Title' } as never,
      ast,
      parser
    );
    const partXml = parser.getZipAdapter().readText('docProps/core.xml');
    expect(partXml).toContain('<cp:lastPrinted>3</cp:lastPrinted>');
    expect(partXml).toContain('Draft');
    expect(partXml).toContain('keepme');
    const iRev = partXml.indexOf('cp:revision');
    const iLp = partXml.indexOf('cp:lastPrinted');
    const iCreated = partXml.indexOf('dcterms:created');
    const iCat = partXml.indexOf('cp:category');
    const iStatus = partXml.indexOf('cp:contentStatus');
    expect(iRev).toBeGreaterThanOrEqual(0);
    expect(iLp).toBeGreaterThanOrEqual(0);
    expect(iCreated).toBeGreaterThanOrEqual(0);
    expect(iRev).toBeLessThan(iLp);
    expect(iLp).toBeLessThan(iCreated);
    expect(iCat).toBeGreaterThanOrEqual(0);
    expect(iStatus).toBeGreaterThanOrEqual(0);
    expect(iCat).toBeLessThan(iStatus);
  });

  it('T6 non-integer revision rejects loudly (no silent Math.trunc)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    expect(() =>
      new CorePropertiesManipulator(ast).update({ revision: 7.5 } as never)
    ).toThrow(/E_VALIDATION/);
    expect(() => mergeCorePropertiesXml(null, { revision: 7.5 } as never)).toThrow(/E_VALIDATION/);
    const badRev =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
      `<cp:revision>7.5</cp:revision></cp:coreProperties>`;
    expect(() => parseCorePropertiesXml(badRev)).toThrow(/E_VALIDATION/);
  });

  it('T7 mergeCorePropertiesXml XXE/DOCTYPE check is loud', () => {
    expect(() =>
      mergeCorePropertiesXml(
        `<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY xxe "evil">]><cp:coreProperties></cp:coreProperties>`,
        { title: 'x' }
      )
    ).toThrow(/XXE|DTD|Security/);
    expect(() => parseCorePropertiesXml(`<!DOCTYPE foo><cp:coreProperties/>`)).toThrow(
      /XXE|DTD|Security/
    );
  });

  it('T8a revision-via-op rejects (strict schema + executor read-only guard)', async () => {
    expect(() =>
      validateOperationsStrict([{ action: 'metadata', operation: 'setProperties', revision: 7 }])
    ).toThrow(/E_VALIDATION/);
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    await expect(
      executeMetadata(
        { action: 'metadata', operation: 'setProperties', revision: 7 } as never,
        ast,
        parser
      )
    ).rejects.toThrow(/E_VALIDATION/);
  });

  it('T8b missing operation rejects (schema-aligned, no op fallback)', async () => {
    expect(() => validateOperationsStrict([{ action: 'metadata', title: 'NoOp' }])).toThrow(
      /E_VALIDATION/
    );
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    await expect(
      executeMetadata({ action: 'metadata', title: 'NoOp' } as never, ast, parser)
    ).rejects.toThrow(/E_VALIDATION/);
  });

  it('T9 windowed inspect (offset/limit + aroundNode) still surfaces metadata.title', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'e2-window.docx');
    const batch = await handleExecuteBatch(store, {
      sessionId: handle.sessionId,
      operations: [
        { action: 'metadata', operation: 'setProperties', title: 'Window Title' },
      ],
      sequential: true,
    });
    expect(batch).toMatchObject({ ok: true });
    const paged = await handleInspectDocument(store, {
      sessionId: handle.sessionId,
      fidelity: 'standard',
      offset: 0,
      limit: 1,
    });
    expect(paged).toMatchObject({ ok: true });
    if (paged.ok !== true) throw new Error('paged inspect failed');
    expect(paged.metadata.title).toBe('Window Title');
    const anchorId = handle.ast.sections[0]?.nodes[0]?.id;
    expect(anchorId).toBeDefined();
    const zoomed = await handleInspectDocument(store, {
      sessionId: handle.sessionId,
      fidelity: 'standard',
      aroundNode: anchorId as string,
      window: 1,
    });
    expect(zoomed).toMatchObject({ ok: true });
    if (zoomed.ok !== true) throw new Error('zoomed inspect failed');
    expect(zoomed.metadata.title).toBe('Window Title');
  });
});

describe('new documents carry their own properties', () => {
  it.each(['blank', 'report', 'memo'])('a fresh %s document has title, author and real times', async (template) => {
    const before = Date.now();
    const parser = new DocxParser(createBlankDocxBuffer(template));
    const ast = await parser.getAST();
    expect(ast.metadata.title).toBeTruthy();
    expect(ast.metadata.creator).toBe('ASTra');
    expect(ast.metadata.lastModifiedBy).toBe('ASTra');
    const created = new Date(String(ast.metadata.created)).getTime();
    expect(created).toBeGreaterThan(before - 5000);
    expect(created).toBeLessThanOrEqual(Date.now() + 1000);
  });

  it('carries settings, a font table and a theme, all registered', () => {
    const zip = new AdmZip(createBlankDocxBuffer('report'));
    for (const part of ['word/settings.xml', 'word/fontTable.xml', 'word/theme/theme1.xml', 'docProps/app.xml']) {
      expect(zip.getEntry(part), part).not.toBeNull();
      expect(zip.readAsText('[Content_Types].xml')).toContain(`/${part}`);
    }
    const rels = zip.readAsText('word/_rels/document.xml.rels');
    for (const target of ['settings.xml', 'fontTable.xml', 'theme/theme1.xml']) expect(rels).toContain(`Target="${target}"`);
    expect(zip.readAsText('word/settings.xml')).toContain('w:name="compatibilityMode"');
    expect(zip.readAsText('word/styles.xml')).toContain('w:asciiTheme="minorHAnsi"');
  });

  it('takes the author from ASTRA_AUTHOR and pins the stamps in reproducible mode', async () => {
    const prevAuthor = process.env['ASTRA_AUTHOR'];
    const prevRepro = process.env['ASTRA_REPRODUCIBLE'];
    try {
      process.env['ASTRA_AUTHOR'] = 'Jo Writer';
      process.env['ASTRA_REPRODUCIBLE'] = '1';
      const first = createBlankDocxBuffer('blank');
      const second = createBlankDocxBuffer('blank');
      const core = new AdmZip(first).readAsText('docProps/core.xml');
      expect(core).toContain('<dc:creator>Jo Writer</dc:creator>');
      expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">1980-01-01T00:00:00Z</dcterms:created>');
      expect(first.equals(second)).toBe(true);
    } finally {
      if (prevAuthor === undefined) delete process.env['ASTRA_AUTHOR'];
      else process.env['ASTRA_AUTHOR'] = prevAuthor;
      if (prevRepro === undefined) delete process.env['ASTRA_REPRODUCIBLE'];
      else process.env['ASTRA_REPRODUCIBLE'] = prevRepro;
    }
  });
});
