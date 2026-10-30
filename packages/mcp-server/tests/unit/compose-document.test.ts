import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  AdmZip,
  DocxParser,
  compileMarkdown,
  composeDocument,
  normalizePackage,
  previewNormalizations,
} from '@attentionform/astra-engine';
import { createMcpSessionStore, getMcpAllowedRoots } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { assertPathAllowed, handleComposeDocument } from '../../src/tools/compose-document';

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function assetDirWithPng(name = 'tiny.png'): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-compose-'));
  fs.writeFileSync(path.join(dir, name), Buffer.from(TINY_PNG_BASE64, 'base64'));
  return dir;
}

function documentXmlOf(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): string {
  const handle = store.get(sessionId);
  const buf = handle.parser.rezip();
  return new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

function relsXmlOf(store: ReturnType<typeof createMcpSessionStore>, sessionId: string): string {
  const handle = store.get(sessionId);
  const buf = handle.parser.rezip();
  return new AdmZip(buf).getEntry('word/_rels/document.xml.rels')?.getData().toString('utf8') ?? '';
}

const RICH_MARKDOWN = [
  '# Composition Proof',
  '',
  'Opening paragraph for the composed document.',
  '',
  '## Sections',
  '',
  '- first bullet',
  '- second bullet',
  '',
  '1. first step',
  '2. second step',
  '',
  '| Name | Value |',
  '| --- | --- |',
  '| alpha | 1 |',
  '| beta | 2 |',
  '',
  '---',
  '',
  '```ts',
  'const proof = true;',
  '```',
  '',
  'A note reference[^fn] in running text.',
  '',
  '[^fn]: The footnote body.',
  '',
].join('\n');

async function openBlankSession(): Promise<{ store: ReturnType<typeof createMcpSessionStore>; sessionId: string }> {
  const store = createMcpSessionStore();
  const buf = DocxParser.createNewDocument('blank').rezip();
  const handle = await store.openBuffer(buf, 'memory.docx');
  return { store, sessionId: handle.sessionId };
}

describe('compose_document handler', () => {
  it('composes rich Markdown: created ids, applied counts, validation proof', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, { sessionId, markdown: RICH_MARKDOWN });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`compose failed: ${JSON.stringify(res)}`);
    expect(res.appliedCount).toBe(res.opCount);
    expect(res.opCount).toBeGreaterThan(10);
    expect(res.createdNodeIds.length).toBeGreaterThan(0);
    expect(res.revision).toBe(2);
    expect(res.dirty).toBe(true);
    expect(res.validation.valid).toBe(true);
    expect(res.validation.errorCount).toBe(0);
    expect(res.results.length).toBe(res.opCount);
  });

  it("skins a plain compose with the default guide and keeps the old grey under styleGuide: 'legacy'", async () => {
    const table = '| A | B |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |\n';
    const plain = await openBlankSession();
    const res = await handleComposeDocument(plain.store, { sessionId: plain.sessionId, markdown: table });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error('compose failed');
    expect(res.appliedCount).toBe(res.opCount);
    const xml = documentXmlOf(plain.store, plain.sessionId);
    expect(xml).toContain('w:fill="1E3A8A"');
    expect(xml).not.toContain('D9D9D9');

    const old = await openBlankSession();
    const legacy = await handleComposeDocument(old.store, { sessionId: old.sessionId, markdown: table, styleGuide: 'legacy' });
    expect(legacy.ok).toBe(true);
    if (!legacy.ok) throw new Error('compose failed');
    expect(legacy.appliedCount).toBe(legacy.opCount);
    expect(legacy.opCount).toBeLessThan(res.opCount);
    const oldXml = documentXmlOf(old.store, old.sessionId);
    expect(oldXml).toContain('w:fill="D9D9D9"');
    expect(oldXml).not.toContain('1E3A8A');
    // 'legacy' is for this call only: it is not stored as the session guide.
    expect(old.store.get(old.sessionId).styleGuide).toBeUndefined();
  });

  it('returns the compose warning for a [!NOTE] that is not in a blockquote', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, { sessionId, markdown: '# T\n\n[!NOTE] Loose note.\n' });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error('compose failed');
    expect(res.warnings.some((w) => /\[!NOTE\].*blockquote/.test(w))).toBe(true);
  });

  it('replaces pristine blank seeds so client output holds only composed content', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, { sessionId, markdown: '# Appended\n\nNew body.\n' });
    expect(res.ok).toBe(true);
    const after = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Blank Document' });
    expect(after.ok).toBe(true);
    if (!after.ok) throw new Error('search failed');
    expect(after.matches.length).toBe(0);
    const hint = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Start typing' });
    expect(hint.ok).toBe(true);
    if (!hint.ok) throw new Error('search failed');
    expect(hint.matches.length).toBe(0);
    const appended = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Appended' });
    expect(appended.ok).toBe(true);
    if (!appended.ok) throw new Error('search failed');
    expect(appended.matches.length).toBeGreaterThan(0);
  });

  it('appends after existing content without deleting it once the blank is no longer pristine', async () => {
    const { store, sessionId } = await openBlankSession();
    const seeded = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'User note stays', id: 'user-note' },
      ],
    });
    expect(seeded.ok).toBe(true);
    const res = await handleComposeDocument(store, { sessionId, markdown: '# Appended\n\nNew body.\n' });
    expect(res.ok).toBe(true);
    for (const query of ['Blank Document', 'User note stays', 'Appended']) {
      const found = await LEAN6_HANDLERS.search_content(store, { sessionId, query });
      expect(found.ok).toBe(true);
      if (!found.ok) throw new Error(`search failed for ${query}`);
      expect(found.matches.length).toBeGreaterThan(0);
    }
  });

  it('composed text round-trips through search (headings, list items, table cells, code, footnote)', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, { sessionId, markdown: RICH_MARKDOWN });
    expect(res.ok).toBe(true);
    for (const query of ['Composition Proof', 'first bullet', 'second step', 'alpha', 'proof = true']) {
      const found = await LEAN6_HANDLERS.search_content(store, { sessionId, query });
      expect(found.ok).toBe(true);
      if (!found.ok) throw new Error(`search failed for ${query}`);
      expect(found.matches.length).toBeGreaterThan(0);
    }
  });

  it('rejects out-of-scope input fail-closed with a line number and writes nothing', async () => {
    const { store, sessionId } = await openBlankSession();
    const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(before.ok).toBe(true);
    const res = await handleComposeDocument(store, { sessionId, markdown: '# Ok\n\n![](image.png)\n' });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('out-of-scope markdown unexpectedly composed');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('line 3');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok || !before.ok) throw new Error('inspect failed');
    expect(after.toonText).toBe(before.toonText);
    expect(after.revision).toBe(before.revision);
  });

  it('rejects deeper headings, blockquotes and undefined footnotes with line numbers', async () => {
    const { store, sessionId } = await openBlankSession();
    for (const [markdown, line] of [
      ['# Ok\n\n#### Too deep\n', 'line 3'],
      ['> quoted\n', 'line 1'],
      ['Text with a dangling note[^missing].\n', 'line 1'],
    ] as Array<[string, string]>) {
      const res = await handleComposeDocument(store, { sessionId, markdown });
      expect(res.ok).toBe(false);
      if (res.ok) throw new Error(`unexpectedly composed: ${JSON.stringify(markdown)}`);
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toContain(line);
    }
  });

  it('rejects empty input and schema violations without touching the session', async () => {
    const { store, sessionId } = await openBlankSession();
    const empty = await handleComposeDocument(store, { sessionId, markdown: '   \n  \n' });
    expect(empty.ok).toBe(false);
    if (empty.ok) throw new Error('empty markdown unexpectedly composed');
    expect(empty.code).toBe('E_VALIDATION');
    const missing = await handleComposeDocument(store, { sessionId });
    expect(missing.ok).toBe(false);
    if (missing.ok) throw new Error('missing markdown unexpectedly composed');
    expect(missing.code).toBe('E_VALIDATION');
    const wrongSession = await handleComposeDocument(store, { sessionId: 'sess_nope', markdown: '# Hi\n' });
    expect(wrongSession.ok).toBe(false);
    if (wrongSession.ok) throw new Error('unknown session unexpectedly composed');
    expect(wrongSession.code).toBe('E_SESSION');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok) throw new Error('inspect failed');
    expect(after.revision).toBe(1);
  });

  it('rejects a second compose into the same session via the duplicate-id gate', async () => {
    const { store, sessionId } = await openBlankSession();
    const first = await handleComposeDocument(store, { sessionId, markdown: '# First\n\nBody one.\n' });
    expect(first.ok).toBe(true);
    const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(before.ok).toBe(true);
    const second = await handleComposeDocument(store, { sessionId, markdown: '# Second\n\nBody two.\n' });
    expect(second.ok).toBe(false);
    if (second.ok) throw new Error('second compose unexpectedly succeeded');
    expect(second.code).toBe('E_VALIDATION');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok || !before.ok) throw new Error('inspect failed');
    expect(after.toonText).toBe(before.toonText);
    expect(after.revision).toBe(before.revision);
  });

  it('applies sources past the per-chunk cap as sequential chunks', async () => {
    const { store, sessionId } = await openBlankSession();
    const lines: string[] = [];
    for (let i = 1; i <= 520; i++) lines.push(`Paragraph number ${i} with enough words to be real.`);
    const markdown = lines.join('\n\n') + '\n';
    const res = await handleComposeDocument(store, { sessionId, markdown });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`chunked compose failed: ${JSON.stringify(res)}`);
    expect(res.opCount).toBeGreaterThan(500);
    expect(res.appliedCount).toBe(res.opCount);
    expect(res.validation.valid).toBe(true);
    const found = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Paragraph number 520' });
    expect(found.ok).toBe(true);
    if (!found.ok) throw new Error('search failed');
    expect(found.matches.length).toBeGreaterThan(0);
  });

  it('splits inline spans into runs with real hyperlink targets', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: 'Hello **bold** and *italic* and `code` and [docs](https://example.com) end\n',
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`span compose failed: ${JSON.stringify(res)}`);
    const xml = documentXmlOf(store, sessionId);
    expect(xml).toContain('w:b');
    expect(xml).toContain('w:i');
    expect(xml).toContain('Courier New');
    expect(xml).toContain('w:hyperlink');
    expect(xml).toContain('>bold<');
    expect(xml).toContain('>italic<');
    expect(xml).toContain('>docs<');
    expect(xml).not.toContain('**bold**');
    expect(xml).not.toContain('[docs](https://example.com)');
    expect(relsXmlOf(store, sessionId)).toContain('https://example.com');
  });

  it('embeds block images resolved from allowedDir in source order', async () => {
    const dir = assetDirWithPng();
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: 'Before\n\n![Tiny](tiny.png)\n\nAfter\n',
      allowedDir: dir,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`image compose failed: ${JSON.stringify(res)}`);
    const xml = documentXmlOf(store, sessionId);
    expect(xml).toContain('w:drawing');
    expect(xml.indexOf('<w:t>Before</w:t>')).toBeLessThan(xml.indexOf('<w:t>After</w:t>'));
  });

  it('rejects image refs outside allowedDir without writing', async () => {
    const dir = assetDirWithPng();
    const { store, sessionId } = await openBlankSession();
    const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(before.ok).toBe(true);
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: '![Evil](../../etc/passwd)\n',
      allowedDir: dir,
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('escaping image ref unexpectedly composed');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('outside the allowed directory');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok || !before.ok) throw new Error('inspect failed');
    expect(after.toonText).toBe(before.toonText);
    expect(after.revision).toBe(before.revision);
  });

  it('rejects a template mismatch before any write', async () => {
    const { store, sessionId } = await openBlankSession();
    const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(before.ok).toBe(true);
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: 'Body text here.\n',
      template: 'report',
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('mismatched template unexpectedly composed');
    expect(res.code).toBe('E_VALIDATION');
    expect(res.message).toContain('does not hold the report seed');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok || !before.ok) throw new Error('inspect failed');
    expect(after.toonText).toBe(before.toonText);
    expect(after.revision).toBe(before.revision);
  });

  it('accepts a matching template and replaces the blank seed', async () => {
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: 'Composed body.\n',
      template: 'blank',
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`matching template compose failed: ${JSON.stringify(res)}`);
    const found = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Blank Document' });
    expect(found.ok).toBe(true);
    if (!found.ok) throw new Error('search failed');
    expect(found.matches.length).toBe(0);
    const body = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Composed body' });
    expect(body.ok).toBe(true);
    if (!body.ok) throw new Error('search failed');
    expect(body.matches.length).toBeGreaterThan(0);
  });

  it('resolves compose helpers from the engine entry point', () => {
    expect(typeof compileMarkdown).toBe('function');
    expect(typeof composeDocument).toBe('function');
    expect(typeof normalizePackage).toBe('function');
    expect(typeof previewNormalizations).toBe('function');
    const ops = compileMarkdown('# Entry\n\nBody text.\n');
    expect(Array.isArray(ops)).toBe(true);
    expect(ops.length).toBeGreaterThan(0);
    const preview = previewNormalizations({});
    expect(Array.isArray(preview)).toBe(true);
    const normalized = normalizePackage({});
    expect(normalized).toHaveProperty('parts');
    expect(normalized).toHaveProperty('normalized');
  });

  it('rejects an asset directory outside the workspace roots without writing', async () => {
    const { store, sessionId } = await openBlankSession();
    const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(before.ok).toBe(true);
    const hostile = path.join(path.sep, 'no-such-astra-jail-root-xyz');
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: '# Safe\n\nBody text.\n',
      allowedDir: hostile,
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('hostile asset directory unexpectedly composed');
    expect(res.code).toBe('E_SANDBOX');
    expect(res.message).toContain('outside the allowed directories');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok || !before.ok) throw new Error('inspect failed');
    expect(after.toonText).toBe(before.toonText);
    expect(after.revision).toBe(before.revision);
  });

  it('rejects an existing system directory outside the roots before compiling', async () => {
    const { store, sessionId } = await openBlankSession();
    const hostile = '/etc';
    let directThrew = false;
    try {
      assertPathAllowed(hostile, ['/no-such-astra-jail-root-xyz']);
    } catch (err: unknown) {
      directThrew = true;
      expect((err as { code?: string }).code).toBe('E_SANDBOX');
    }
    expect(directThrew).toBe(true);
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: 'Before\n\n![Tiny](tiny.png)\n\nAfter\n',
      allowedDir: hostile,
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('system asset directory unexpectedly composed');
    expect(res.code).toBe('E_SANDBOX');
  });

  it('accepts an asset directory inside the system temp dir', async () => {
    const dir = assetDirWithPng();
    const { store, sessionId } = await openBlankSession();
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: 'Before\n\n![Tiny](tiny.png)\n\nAfter\n',
      allowedDir: dir,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error(`temp-dir asset compose failed: ${JSON.stringify(res)}`);
    const xml = documentXmlOf(store, sessionId);
    expect(xml).toContain('w:drawing');
  });

  it('rejects allowedDir with literal .. traversal without writing', async () => {
    const dir = assetDirWithPng();
    const traversal = `${dir}${path.sep}..${path.sep}${path.basename(dir)}`;
    expect(traversal.split(/[\\/]/).includes('..')).toBe(true);
    let directThrew = false;
    try {
      assertPathAllowed(traversal, getMcpAllowedRoots());
    } catch (err: unknown) {
      directThrew = true;
      expect((err as { code?: string }).code).toBe('E_SANDBOX');
      expect((err as Error).message).toContain('..');
    }
    expect(directThrew).toBe(true);
    const { store, sessionId } = await openBlankSession();
    const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(before.ok).toBe(true);
    const res = await handleComposeDocument(store, {
      sessionId,
      markdown: '# Safe\n\nBody text.\n',
      allowedDir: traversal,
    });
    expect(res.ok).toBe(false);
    if (res.ok) throw new Error('traversal asset directory unexpectedly composed');
    expect(res.code).toBe('E_SANDBOX');
    expect(res.message).toContain('..');
    const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(after.ok).toBe(true);
    if (!after.ok || !before.ok) throw new Error('inspect failed');
    expect(after.toonText).toBe(before.toonText);
    expect(after.revision).toBe(before.revision);
  });

  it('rejects a symlink inside the roots pointing outside without writing', async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-jail-'));
    const link = path.join(base, 'link-escape');
    try {
      fs.symlinkSync('/etc', link);
    } catch (err: unknown) {
      throw new Error(`symlink setup failed: ${String((err as Error)?.message ?? err)}`);
    }
    try {
      expect(path.resolve(link).startsWith(path.resolve(os.tmpdir()))).toBe(true);
      let directThrew = false;
      try {
        assertPathAllowed(link, getMcpAllowedRoots());
      } catch (err: unknown) {
        directThrew = true;
        expect((err as { code?: string }).code).toBe('E_SANDBOX');
        expect((err as Error).message).toContain('outside the allowed directories');
      }
      expect(directThrew).toBe(true);
      const { store, sessionId } = await openBlankSession();
      const before = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
      expect(before.ok).toBe(true);
      const res = await handleComposeDocument(store, {
        sessionId,
        markdown: '# Safe\n\nBody text.\n',
        allowedDir: link,
      });
      expect(res.ok).toBe(false);
      if (res.ok) throw new Error('symlink-escape asset directory unexpectedly composed');
      expect(res.code).toBe('E_SANDBOX');
      expect(res.message).toContain('outside the allowed directories');
      const after = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
      expect(after.ok).toBe(true);
      if (!after.ok || !before.ok) throw new Error('inspect failed');
      expect(after.toonText).toBe(before.toonText);
      expect(after.revision).toBe(before.revision);
    } finally {
      try {
        fs.rmSync(base, { recursive: true, force: true });
      } catch {
        /* best-effort cleanup */
      }
    }
  });
});
