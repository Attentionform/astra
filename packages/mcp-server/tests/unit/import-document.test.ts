import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleComposeDocument } from '../../src/tools/compose-document';
import { handleImportDocument } from '../../src/lib/mcp/handlers/parity';

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const createdDirs: string[] = [];
afterEach(() => {
  while (createdDirs.length > 0) {
    const dir = createdDirs.pop() as string;
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      /* best-effort cleanup */
    }
  }
});

function scratchDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-import-'));
  createdDirs.push(dir);
  return dir;
}

async function openBlankSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const buf = DocxParser.createNewDocument('blank').rezip();
  const handle = await store.openBuffer(buf, 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

function documentXmlOf(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): string {
  const buf = store.get(sessionId).parser.rezip();
  return new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

const SOURCE = ['# Import Proof', '', 'First body paragraph.', '', '- one', '- two', ''].join('\n');

describe('import_document handler', () => {
  it('ingests markdown byte-equal to the same source composed inline', async () => {
    const dir = scratchDir();
    const mdPath = path.join(dir, 'notes.md');
    fs.writeFileSync(mdPath, SOURCE, 'utf8');
    const viaImport = await openBlankSession();
    const imported = await handleImportDocument(viaImport.store, { sessionId: viaImport.sessionId, path: mdPath });
    expect(imported.ok).toBe(true);
    if (!imported.ok) throw new Error(`import failed: ${JSON.stringify(imported)}`);
    const viaInline = await openBlankSession();
    const inline = await handleComposeDocument(viaInline.store, { sessionId: viaInline.sessionId, markdown: SOURCE });
    expect(inline.ok).toBe(true);
    if (!imported.ok || !inline.ok) throw new Error('compose failed');
    expect(imported.appliedCount).toBe(inline.appliedCount);
    expect(documentXmlOf(viaImport.store, viaImport.sessionId)).toBe(documentXmlOf(viaInline.store, viaInline.sessionId));
    const found = await LEAN6_HANDLERS.search_content(viaImport.store, { sessionId: viaImport.sessionId, query: 'Import Proof' });
    expect(found.ok).toBe(true);
    if (!found.ok) throw new Error('search failed');
    expect(found.matches.length).toBeGreaterThan(0);
  });

  it('resolves image refs under allowedDir', async () => {
    const dir = scratchDir();
    fs.writeFileSync(path.join(dir, 'tiny.png'), Buffer.from(TINY_PNG_BASE64, 'base64'));
    const mdPath = path.join(dir, 'pics.md');
    fs.writeFileSync(mdPath, 'Before\n\n![Tiny](tiny.png)\n\nAfter\n', 'utf8');
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: mdPath, allowedDir: dir });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`import failed: ${JSON.stringify(res)}`);
    expect(documentXmlOf(store, sessionId)).toContain('w:drawing');
  });

  it('refuses html with the cli-identical verdict', async () => {
    const dir = scratchDir();
    const htmlPath = path.join(dir, 'page.html');
    fs.writeFileSync(htmlPath, '<h1>Hi</h1>', 'utf8');
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: htmlPath });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('html unexpectedly imported');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('HTML import is not supported directly');
    expect(res.remedy).toContain('Convert the HTML file to Markdown');
  });

  it('refuses docx naming open_document', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: path.join(scratchDir(), 'report.docx') });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('docx unexpectedly imported');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('open_document');
    expect(res.remedy).toContain('open_document');
  });

  it('refuses outside-jail reads with E_SANDBOX', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: '/etc/x.md' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('outside-jail read unexpectedly imported');
    expect(res.code).toBe('E_SANDBOX');
    expect(res.remedy.length).toBeGreaterThan(0);
  });

  it('refuses missing files with E_IO', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: path.join(scratchDir(), 'absent.md') });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('missing file unexpectedly imported');
    expect(res.code).toBe('E_IO');
    expect(res.remedy.length).toBeGreaterThan(0);
  });

  it('refuses oversize files with E_LIMIT', async () => {
    const dir = scratchDir();
    const bigPath = path.join(dir, 'big.md');
    fs.writeFileSync(bigPath, `# Big\n\n${'x '.repeat(300000)}\n`, 'utf8');
    expect(fs.statSync(bigPath).size).toBeGreaterThan(500000);
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: bigPath });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('oversize file unexpectedly imported');
    expect(res.code).toBe('E_LIMIT');
    expect(res.remedy.length).toBeGreaterThan(0);
  });

  it('refuses empty files with E_VALIDATION', async () => {
    const dir = scratchDir();
    const emptyPath = path.join(dir, 'empty.md');
    fs.writeFileSync(emptyPath, '', 'utf8');
    const { store, sessionId } = await openBlankSession();
    const res = await handleImportDocument(store, { sessionId, path: emptyPath });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('empty file unexpectedly imported');
    expect(res.code).toBe('E_VALIDATION');
  });

  it('requires a session id (never auto-creates)', async () => {
    const dir = scratchDir();
    const mdPath = path.join(dir, 'notes.md');
    fs.writeFileSync(mdPath, SOURCE, 'utf8');
    const { store } = await openBlankSession();
    const missing = await handleImportDocument(store, { path: mdPath });
    expect(missing.ok).toBe(false);
    if (missing.ok) throw new Error('sessionless import unexpectedly succeeded');
    expect(missing.code).toBe('E_VALIDATION');
  });

  it('refuses an unknown session with E_SESSION', async () => {
    const dir = scratchDir();
    const mdPath = path.join(dir, 'notes.md');
    fs.writeFileSync(mdPath, SOURCE, 'utf8');
    const store = createMcpSessionStore();
    const res = await handleImportDocument(store, { sessionId: 'sess_nope', path: mdPath });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('unknown session unexpectedly imported');
    expect(res.code).toBe('E_SESSION');
  });
});
