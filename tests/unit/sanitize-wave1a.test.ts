import { describe, it, expect, afterAll } from 'vitest';
import AdmZip from 'adm-zip';
import { createHash } from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentAST,
  DocxParser,
  OperationEngine,
  SANITIZER_DEFAULT_AUTHOR,
  createBlankDocxBuffer,
  executeHyperlink,
  executeImage,
  isBlockedExecutableEntry,
  sanitizeAppPropertiesXml,
  sanitizeCorePropertiesXml,
  sanitizeDocxBuffer,
  sanitizeDocumentXml,
  sanitizePackageEntries,
  resolveDeterministicZipDate,
  toDosTimeUtc,
  validateDocxBuffer,
} from '@attentionform/astra-engine';
// Central XXE gate: packages/engine/src/ooxml/xxe-guard.ts:83 (assertNoXxe;
// alias assertNoForbiddenMarkup at :90, re-exported via sanitizer.ts:7).
// containsForbiddenMarkup (xxe-guard.ts:61) is NOT re-exported from the engine
// index, so the stale engine import is replaced with this direct import
// (precedent: tests/unit/xxe-guard-wave.test.ts:30-33).
import { assertNoXxe } from '../../packages/engine/src/ooxml/xxe-guard';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function createMockAST(): DocumentAST {
  const doc = new DOMParser().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>anchor</w:t></w:r></w:p></w:body></w:document>',
    'text/xml'
  );
  const bodyP = doc.getElementsByTagName('w:p')[0];
  return {
    metadata: {},
    sections: [
      {
        id: 'sec_root',
        heading: null,
        nodes: [
          {
            id: 'p_001',
            type: 'paragraph',
            xml: bodyP,
            metadata: {
              text: 'anchor',
              hasCitation: false,
              hasComment: false,
              hasBookmark: false,
              hasTrackedChanges: false,
              hasEquation: false,
              hasField: false,
              hasImage: false,
              hasHyperlink: false,
            },
            version: 1,
            checksum: 'chk_1',
          },
        ],
        pageSetup: {},
        version: 1,
        checksum: 'sec_chk_1',
      },
    ],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1,
  };
}

function docXmlWithBodyText(inner: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>` +
    `<w:p><w:r><w:t>${inner}</w:t></w:r></w:p>` +
    `</w:body></w:document>`
  );
}

// There is no exported standalone hyperlink URL validator in the engine
// (verified: 'assertSafeHyperlinkUrl' not present in packages/engine/src/index;
// the scheme gate lives inline in packages/engine/src/operations/hyperlink.ts:119-142
// and is exercised through executeHyperlink). These helpers keep the original
// intent (dangerous schemes blocked, safe shapes pass) against the real API.
function expectLinkBlocked(url: string): void {
  expect(() =>
    executeHyperlink({ action: 'hyperlink', target: 'p_001', text: 'x', url } as any, createMockAST())
  ).toThrow(/E_VALIDATION/);
}

function expectLinkAllowed(url: string): void {
  expect(() =>
    executeHyperlink({ action: 'hyperlink', target: 'p_001', text: 'x', url } as any, createMockAST())
  ).not.toThrow();
}

describe('1A hardening: image path jail, sanitize ingress, body PII, hyperlink schemes', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-1a-'));
  const evilPng = path.join(tmpDir, 'evil.png');
  fs.writeFileSync(evilPng, Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
  fs.mkdirSync(path.join(tmpDir, 'sub'), { recursive: true });

  afterAll(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('image.fileData rejects an absolute /tmp file path with E_VALIDATION + data-URI remedy', () => {
    const ast = createMockAST();
    expect(() =>
      executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: evilPng } as any, ast)
    ).toThrow(/E_VALIDATION/);
    // Exact engine text (packages/engine/src/operations/image.ts:551-554,
    // verified by live call): filesystem paths fail the base64 alphabet check.
    expect(() =>
      executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: evilPng } as any, createMockAST())
    ).toThrow(/Invalid image data string/);
    expect(() =>
      executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: evilPng } as any, createMockAST())
    ).toThrow(/Provide image as data: URL or base64 string/);
    // Rejected before any media lands in the AST.
    expect(ast.sections[0].nodes.length).toBe(1);
  });

  it('image.fileData rejects relative and traversal paths resolving to existing files', () => {
    const rel = path.relative(process.cwd(), evilPng);
    expect(path.isAbsolute(rel)).toBe(false);
    expect(fs.existsSync(rel)).toBe(true);
    expect(() =>
      executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: rel } as any, createMockAST())
    ).toThrow(/E_VALIDATION/);

    const traversal = tmpDir + path.sep + 'sub' + path.sep + '..' + path.sep + 'evil.png';
    expect(traversal).toContain('..');
    expect(fs.existsSync(traversal)).toBe(true);
    expect(() =>
      executeImage(
        { action: 'image', operation: 'insert', anchor: 'p_001', fileData: traversal } as any,
        createMockAST()
      )
    ).toThrow(/E_VALIDATION/);
  });

  it('image.fileData still accepts data: URIs and raw base64', () => {
    const ast1 = createMockAST();
    executeImage(
      { action: 'image', operation: 'insert', anchor: 'p_001', fileData: `data:image/png;base64,${PNG_1PX}` } as any,
      ast1
    );
    expect(ast1.sections[0].nodes.length).toBe(2);
    const ast2 = createMockAST();
    executeImage({ action: 'image', operation: 'insert', anchor: 'p_001', fileData: PNG_1PX } as any, ast2);
    expect(ast2.sections[0].nodes.length).toBe(2);
  });

  it('sanitize ingress rejects an XXE part fail-closed (E_VALIDATION)', () => {
    const xxe =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>` +
      `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p></w:body></w:document>`;
    expect(() => sanitizeDocumentXml(xxe, {})).toThrow(/E_VALIDATION/);
    const files = new Map<string, Buffer>([['word/document.xml', Buffer.from(xxe, 'utf8')]]);
    expect(() => sanitizePackageEntries(files, {})).toThrow(/E_VALIDATION/);
  });

  it('sanitize ingress rejects lowercase doctype/entity variants (case-insensitive gate)', () => {
    // Central guard packages/engine/src/ooxml/xxe-guard.ts:83; pattern at :58
    // is case-insensitive over DOCTYPE/ENTITY. Live-verified: lowercase
    // '<!doctype' throws E_VALIDATION with a remedy.
    for (const decl of [
      '<!DOCTYPE foo>',
      '<!doctype foo>',
      '<!DocType foo>',
      '<!ENTITY xxe "evil">',
      '<!entity xxe "evil">',
    ]) {
      expect(() => assertNoXxe(decl, 'probe'), decl).toThrow(/E_VALIDATION/);
    }
    expect(() =>
      assertNoXxe(`<w:p xmlns:w="${W_NS}"><w:r><w:t>clean</w:t></w:r></w:p>`, 'probe')
    ).not.toThrow();
    const lower =
      `<?xml version="1.0"?>` +
      `<!doctype foo [<!entity xxe "evil">]>` +
      `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`;
    expect(() => sanitizeDocumentXml(lower, {})).toThrow(/E_VALIDATION/);
    const files = new Map<string, Buffer>([['word/document.xml', Buffer.from(lower, 'utf8')]]);
    expect(() => sanitizePackageEntries(files, {})).toThrow(/E_VALIDATION/);
  });

  it('sanitize ingress drops rogue zip entries and denied executables with names+count in report', () => {
    // sanitizer.ts isRogueZipEntry covers traversal / absolute / drive /
    // control shapes; isBlockedExecutableEntry covers exe dll bat cmd ps1
    // vbs js jse scr msi (case-insensitive, final segment). Both take the
    // same silent drop+count path into report.droppedRogue (never a throw,
    // never echoed bytes) so hostile entries cannot break legit export.
    expect(isBlockedExecutableEntry('evil/malware.exe')).toBe(true);
    expect(isBlockedExecutableEntry('evil/TOOL.BAT')).toBe(true);
    expect(isBlockedExecutableEntry('evil/nested/run.PS1')).toBe(true);
    expect(isBlockedExecutableEntry('word/document.xml')).toBe(false);
    expect(isBlockedExecutableEntry('word/media/image.exe.png')).toBe(false);
    const good = docXmlWithBodyText('hello');
    const typesXml = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
    const relsXml = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
    const coreXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>`;
    const oleBytes = Buffer.from([0xd0, 0xcf, 0x11, 0xe0]);
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from(good, 'utf8')],
      ['[Content_Types].xml', Buffer.from(typesXml, 'utf8')],
      ['word/_rels/document.xml.rels', Buffer.from(relsXml, 'utf8')],
      ['docProps/core.xml', Buffer.from(coreXml, 'utf8')],
      // M4: embeddings bins refuse by default (opaque macro surface), so the
      // legit-opaque control lives outside embeddings (R-10 corroboration
      // rule keeps unreferenced non-embeddings bins legit).
      ['word/oleObject1.bin', oleBytes],
      ['../evil.xml', Buffer.from('rogue', 'utf8')],
      ['/abs/path.xml', Buffer.from('rogue', 'utf8')],
      ['evil/malware.exe', Buffer.from('MZ', 'utf8')],
      ['evil/TOOL.BAT', Buffer.from('rogue', 'utf8')],
    ]);
    expect(() => sanitizePackageEntries(files, {})).not.toThrow();
    const res = sanitizePackageEntries(files, {});
    const out: Map<string, Buffer> = res.files;
    // Hostile entries gone from output, listed by name with an honest count.
    for (const name of ['../evil.xml', '/abs/path.xml', 'evil/malware.exe', 'evil/TOOL.BAT']) {
      expect(out.has(name), name).toBe(false);
      expect(res.report.droppedRogue, name).toContain(name);
    }
    expect(res.report.droppedRogue.length).toBe(4);
    // Legit parts (word/, _rels, [Content_Types], docProps, opaque bin)
    // pass through byte-identical.
    expect(out.get('word/document.xml')!.toString('utf8')).toContain('hello');
    expect(out.get('[Content_Types].xml')!.toString('utf8')).toBe(typesXml);
    expect(out.get('word/_rels/document.xml.rels')!.toString('utf8')).toBe(relsXml);
    expect(out.get('docProps/core.xml')!.toString('utf8')).toBe(coreXml);
    expect(out.get('word/oleObject1.bin')).toEqual(oleBytes);
    // Canonical names still pass with an empty drop list.
    const ok = new Map<string, Buffer>([['word/document.xml', Buffer.from(good, 'utf8')]]);
    expect(sanitizePackageEntries(ok, {}).report.droppedRogue).toEqual([]);
  });

  it('sanitizeDocxBuffer drops rogue+exe entries, keeps legit parts, placeholder core stays valid OOXML', async () => {
    // Buffer-level drop path (sanitizer.ts sanitizeDocxBuffer): rogue and
    // denied-executable entries never reach the package map, legit parts
    // survive, and the zero-retention '' placeholders keep schema validity.
    // Note: adm-zip normalizes ../ and leading / on write, so the hostile
    // names here use shapes that survive the zip layer (drive letter for
    // rogue, exe extension for the deny-list); ../ and /abs shapes are
    // pinned at the package-map level in the test above.
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.addFile('D:/evil.txt', Buffer.from('rogue', 'utf8'));
    zip.addFile('evil/malware.exe', Buffer.from('MZ', 'utf8'));
    const { buffer, report } = sanitizeDocxBuffer(zip.toBuffer(), { zeroRetention: true });
    expect(report.ok).toBe(true);
    expect(report.droppedRogue).toContain('D:/evil.txt');
    expect(report.droppedRogue).toContain('evil/malware.exe');
    expect(report.droppedRogue.length).toBe(2);
    const outZip = new AdmZip(buffer);
    expect(outZip.getEntry('D:/evil.txt')).toBeNull();
    expect(outZip.getEntry('evil/malware.exe')).toBeNull();
    expect(outZip.getEntry('word/document.xml')).not.toBeNull();
    expect(outZip.getEntry('[Content_Types].xml')).not.toBeNull();
    const v = await validateDocxBuffer(buffer);
    expect(v.ok).toBe(true);
  });

  it('body SSN in w:t is flagged: bodyPII > 0 and text redacted (never silent pii=0)', () => {
    // Real field is bodyPII (packages/engine/src/sanitizer.ts:421; return
    // object :501; report aggregates at :1244-1245). There is no purgedBodyPII
    // key (verified by printing Object.keys of a live result).
    const xml = docXmlWithBodyText('SSN 123-45-6789 on file');
    const r = sanitizeDocumentXml(xml, {});
    expect(r.bodyPII).toBe(1);
    expect(r.xml).not.toContain('123-45-6789');
    expect(r.xml).toContain('[redacted]');
    const files = new Map<string, Buffer>([['word/document.xml', Buffer.from(xml, 'utf8')]]);
    const { files: clean, report } = sanitizePackageEntries(files, {});
    expect(report.purgedPII).toBeGreaterThanOrEqual(1);
    expect(clean.get('word/document.xml')!.toString('utf8')).not.toContain('123-45-6789');
  });

  it('clean body text reports bodyPII 0 and stays byte-identical', () => {
    const xml = docXmlWithBodyText('Hello world');
    const r = sanitizeDocumentXml(xml, {});
    expect(r.bodyPII).toBe(0);
    expect(r.xml).toBe(xml);
  });

  it('core keywords/title/subject are purged under zero-retention and kept otherwise', () => {
    const core =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">` +
      `<dc:title>Secret Project Falcon</dc:title>` +
      `<dc:subject>Stealth merger notes</dc:subject>` +
      `<cp:keywords>falcon, stealth, q3-layoffs</cp:keywords>` +
      `</cp:coreProperties>`;
    const purged = sanitizeCorePropertiesXml(core, { zeroRetention: true });
    expect(purged.xml).not.toContain('Secret Project Falcon');
    expect(purged.xml).not.toContain('Stealth merger notes');
    expect(purged.xml).not.toContain('falcon, stealth');
    // Count covers the metadata purge too: creator/modifier replacements
    // plus each zero-retention tag emptied to the '' placeholder (title,
    // subject, keywords present here = 3). Placeholder strategy unchanged —
    // no attribution string survives, and the empty elements stay valid.
    expect(purged.purgedAuthors).toBeGreaterThanOrEqual(3);
    const kept = sanitizeCorePropertiesXml(core, { zeroRetention: false });
    expect(kept.xml).toContain('Secret Project Falcon');
    expect(kept.xml).toContain('falcon, stealth');
  });

  it('hyperlink rejects javascript:/data:text/html/vbscript: with E_VALIDATION', () => {
    // No standalone assertSafeHyperlinkUrl export exists in the engine; the
    // scheme allow-list lives in operations/hyperlink.ts:119-142 and throws
    // "E_VALIDATION: Invalid hyperlink URL scheme '<scheme>'…" (live-verified).
    for (const url of [
      'javascript:alert(1)',
      'JAVASCRIPT:alert(1)',
      'data:text/html,<script>alert(1)</script>',
      'vbscript:msgbox(1)',
      'VBScript:msgbox(1)',
    ]) {
      expectLinkBlocked(url);
    }
    expect(() =>
      executeHyperlink(
        { action: 'hyperlink', target: 'p_001', text: 'x', url: 'javascript:alert(1)' } as any,
        createMockAST()
      )
    ).toThrow(/E_VALIDATION/);
  });

  it('hyperlink allows http/https/mailto plus relative anchors', () => {
    for (const url of [
      'https://example.com',
      'http://example.com/a?b=c',
      'HTTPS://EXAMPLE.COM',
      'mailto:jane@example.com',
      '#bm_1',
      'docs/guide',
      '/root/relative',
    ]) {
      expectLinkAllowed(url);
    }
    const ast = createMockAST();
    executeHyperlink(
      { action: 'hyperlink', target: 'p_001', text: 'site', url: 'https://example.com' } as any,
      ast
    );
    expect(ast.sections[0].nodes[0].xml.getElementsByTagName('w:hyperlink').length).toBe(1);
  });

  it('sanitizeDocxBuffer is deterministic: double sanitize byte-identical with clamped entry times', async () => {
    // sanitizer.ts sanitizeDocxBuffer clamps via zip-hardening clampZipEntryTimes
    // (same default epoch as parser rezip/create paths), so identical inputs
    // hash-identically instead of diverging on adm-zip wall-clock mtimes.
    const src = createBlankDocxBuffer('blank');
    const a = sanitizeDocxBuffer(src, {});
    const b = sanitizeDocxBuffer(src, {});
    expect(a.buffer.equals(b.buffer)).toBe(true);
    const ha = createHash('sha256').update(a.buffer).digest('hex');
    const hb = createHash('sha256').update(b.buffer).digest('hex');
    expect(ha).toBe(hb);
    const v = await validateDocxBuffer(a.buffer);
    expect(v.ok).toBe(true);
    // DOS 2s granularity can mask a missing clamp on rapid double writes, so
    // pin the clamp directly: every entry timeval equals the deterministic epoch.
    const expected = toDosTimeUtc(resolveDeterministicZipDate());
    const entries = new AdmZip(a.buffer).getEntries();
    expect(entries.length).toBeGreaterThan(0);
    for (const e of entries) {
      expect((e.header as unknown as { timeval: number }).timeval).toBe(expected);
    }
  });
});

describe('R-10 batch/save ingress drops rogue entries with droppedRogue accounting', () => {
  const ROGUE_NAMES = ['D:/evil.txt', 'evil/malware.exe'];
  const OLE_BYTES = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0x01, 0x02]);

  function rogueDocxBuffer(): Buffer {
    // Drive-letter (rogue shape) + exe (deny-list) survive the zip layer
    // (same precedent as the sanitizeDocxBuffer test above); ../ and /abs
    // shapes are pinned at the package-map level there.
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    // M4: embeddings bins refuse by default — the legit-opaque control lives
    // outside embeddings (unreferenced non-embeddings bins stay legit).
    zip.addFile('word/oleObject1.bin', OLE_BYTES);
    for (const name of ROGUE_NAMES) zip.addFile(name, Buffer.from('rogue', 'utf8'));
    return zip.toBuffer();
  }

  it('rogue-entry docx through batch: rogues dropped + counted, legit bin byte-identical, output validates', async () => {
    const parser = new DocxParser(rogueDocxBuffer());
    // Parse ingress drops + counts via the shared sanitizer predicates
    // (same drop+count policy as sanitize report.droppedRogue).
    expect(parser.getDroppedRogueEntries()).toEqual([...ROGUE_NAMES].sort());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor, position: 'before', text: 'R10 legit edit' } as any,
      ],
      sequential: true,
    });
    expect(res.success).toBe(true);
    // Save ingress keeps the drop + count across the batch round-trip.
    const out = parser.rezip();
    expect(parser.getDroppedRogueEntries()).toEqual([...ROGUE_NAMES].sort());
    expect(parser.getDroppedRogueEntries()).toHaveLength(2);
    const outZip = new AdmZip(out);
    for (const name of ROGUE_NAMES) expect(outZip.getEntry(name), name).toBeNull();
    expect(outZip.getEntry('word/oleObject1.bin')!.getData()).toEqual(OLE_BYTES);
    expect(outZip.getEntry('word/document.xml')!.getData().toString('utf8')).toContain('R10 legit edit');
    const v = await validateDocxBuffer(out);
    expect(v.ok).toBe(true);
  });

  it('clean docx through batch reports empty droppedRogue and validates', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    expect(parser.getDroppedRogueEntries()).toEqual([]);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor, position: 'before', text: 'clean edit' } as any,
      ],
      sequential: true,
    });
    expect(res.success).toBe(true);
    const out = parser.rezip();
    expect(parser.getDroppedRogueEntries()).toEqual([]);
    const v = await validateDocxBuffer(out);
    expect(v.ok).toBe(true);
  });
});

describe('sanitized properties carry no vendor fingerprint', () => {
  it('default replacement is empty', () => {
    expect(SANITIZER_DEFAULT_AUTHOR).toBe('');
  });

  it('bare sanitize empties company and manager without vendor bytes', () => {
    const app =
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
      `<Company>Acme Corp</Company><Manager>Jane Doe</Manager></Properties>`;
    const r = sanitizeAppPropertiesXml(app, {});
    expect(r.purgedAuthors).toBe(2);
    expect(r.xml).not.toContain('Acme Corp');
    expect(r.xml).not.toContain('Jane Doe');
    expect(r.xml).not.toContain('ASTra');
    expect(r.xml).toContain('<Company></Company>');
    expect(r.xml).toContain('<Manager></Manager>');
  });

  it('bare sanitize empties creator fields without vendor bytes', () => {
    const core =
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
      `<dc:creator>Alice Doe</dc:creator>` +
      `<cp:lastModifiedBy>Bob Smith</cp:lastModifiedBy>` +
      `</cp:coreProperties>`;
    const r = sanitizeCorePropertiesXml(core, {});
    expect(r.xml).not.toContain('Alice Doe');
    expect(r.xml).not.toContain('Bob Smith');
    expect(r.xml).not.toContain('ASTra');
    expect(r.xml).toContain('<dc:creator></dc:creator>');
    expect(r.xml).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');
  });

  it('bare sanitize clears author text but keeps a single-letter placeholder', () => {
    // Empty default leaves no name bytes; the initials field still carries
    // a one-letter fallback derived from the empty value.
    const doc =
      `<w:p xmlns:w="${W_NS}" w:author="Jane Doe" w:initials="JD"><w:r><w:t>x</w:t></w:r></w:p>`;
    const r = sanitizeDocumentXml(doc, {});
    expect(r.xml).toContain('w:author=""');
    expect(r.xml).toContain('w:initials="A"');
    expect(r.xml).not.toContain('Jane Doe');
  });

  it('explicit replacement name still stamps core and app properties', () => {
    const core =
      `<cp:coreProperties xmlns:cp="x" xmlns:dc="y">` +
      `<dc:creator>Alice Doe</dc:creator>` +
      `<cp:lastModifiedBy>Bob Smith</cp:lastModifiedBy>` +
      `</cp:coreProperties>`;
    const rc = sanitizeCorePropertiesXml(core, { replacementAuthor: 'CustomName' });
    expect(rc.xml).toContain('CustomName');
    expect(rc.xml).not.toContain('Alice Doe');
    expect(rc.xml).not.toContain('ASTra');
    const app =
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
      `<Company>Acme Corp</Company><Manager>Jane Doe</Manager></Properties>`;
    const ra = sanitizeAppPropertiesXml(app, { replacementAuthor: 'CustomName' });
    expect(ra.xml).toContain('CustomName');
    expect(ra.xml).not.toContain('Acme Corp');
    expect(ra.xml).not.toContain('ASTra');
    const doc =
      `<w:p xmlns:w="${W_NS}" w:author="Evil Author"><w:r><w:t>x</w:t></w:r></w:p>`;
    const rd = sanitizeDocumentXml(doc, { replacementAuthor: 'CustomName' });
    expect(rd.xml).toContain('w:author="CustomName"');
    expect(rd.xml).not.toContain('Evil Author');
  });

  it('end-to-end bare sanitize leaves docProps free of vendor bytes and validates', async () => {
    const blank = createBlankDocxBuffer('blank');
    const zip = new AdmZip(blank);
    let d = zip.getEntry('word/document.xml')!.getData().toString('utf8');
    d = d.replace(
      'New Document',
      'Hello</w:t></w:r>' +
        '<w:ins w:id="0" w:author="Jane Doe" w:date="2026-01-01T00:00:00Z">' +
        '<w:r><w:t> sneaky</w:t></w:r></w:ins><w:r><w:t> tail'
    );
    zip.updateFile('word/document.xml', Buffer.from(d, 'utf8'));
    zip.addFile(
      'docProps/core.xml',
      Buffer.from(
        `<?xml version="1.0" encoding="UTF-8"?>` +
          `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
          `xmlns:dc="http://purl.org/dc/elements/1.1/">` +
          `<dc:creator>Jane Doe</dc:creator>` +
          `<cp:lastModifiedBy>Jane Doe</cp:lastModifiedBy>` +
          `</cp:coreProperties>`,
        'utf8'
      )
    );
    zip.addFile(
      'docProps/app.xml',
      Buffer.from(
        `<?xml version="1.0" encoding="UTF-8"?>` +
          `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
          `<Company>Acme Corp</Company><Manager>Jane Doe</Manager>` +
          `</Properties>`,
        'utf8'
      )
    );
    const { buffer, report } = sanitizeDocxBuffer(zip.toBuffer(), {});
    expect(report.purgedAuthors).toBeGreaterThanOrEqual(4);
    const out = new AdmZip(buffer);
    for (const name of ['docProps/core.xml', 'docProps/app.xml', 'word/document.xml']) {
      const text = out.getEntry(name)!.getData().toString('utf8');
      expect(text, name).not.toContain('Jane Doe');
      expect(text, name).not.toContain('Acme Corp');
      expect(text, name).not.toContain('ASTra');
    }
    const v = await validateDocxBuffer(buffer);
    expect(v.ok).toBe(true);
  });
});
