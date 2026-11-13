import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { createHash } from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, createBlankDocxBuffer, validateDocxBuffer } from '../../packages/engine/src';
import { setPdfBackendBinaryForTesting, resetPdfBackendCache } from '../../packages/engine/src/export/pdf';
import { run as runCli, EXIT_FAIL, EXIT_OK } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleOpenDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleSaveDocument } from '../../packages/mcp-server/src/lib/mcp/handlers/save-close';
import { handleExportDocument } from '../../packages/mcp-server/src/tools/export';

/**
 * What the PDF backend receives, through the CLI and through MCP. A fake backend copies its input docx aside and
 * writes a stub PDF, so no LibreOffice is needed. Needs the native validator (CI, or OOXML_VALIDATOR_CLI).
 *
 * CLI contract: a valid file is rendered as the original bytes; a file whose only errors are element order is
 * rendered from a verified temporary copy and the source file is never written; anything else is refused before
 * the backend starts.
 * MCP contract (different on purpose): the session is serialized exactly as a save would write it, then validated.
 */
const FIXTURE = path.resolve(__dirname, '../fixtures/word-order.docx');
const GALLERY = path.resolve(__dirname, '../../test-gallery');
const sha = (b: Buffer): string => createHash('sha256').update(b).digest('hex');

let dir: string;
let captured: string;
let marker: string;
let prevTmp: string | undefined;
let prevRoots: string | undefined;
let logged: string[];

beforeEach(() => {
  dir = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'astra-pdf-repair-')));
  prevTmp = process.env.ASTRA_TMPDIR;
  process.env.ASTRA_TMPDIR = path.join(dir, 'tmp');
  captured = path.join(dir, 'captured.docx');
  marker = path.join(dir, 'backend-ran');
  const fake = path.join(dir, 'fake-office');
  fs.writeFileSync(
    fake,
    `#!/bin/sh\ntouch "${marker}"\ncp "$5" "${captured}"\nprintf '%%PDF-1.4 stub' > "\${5%.docx}.pdf"\nexit 0\n`,
    { mode: 0o755 }
  );
  setPdfBackendBinaryForTesting(fake);
  prevRoots = process.env.ASTRA_ALLOWED_DIRS;
  process.env.ASTRA_ALLOWED_DIRS = dir;
  logged = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void logged.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void logged.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  resetPdfBackendCache();
  if (prevRoots === undefined) delete process.env.ASTRA_ALLOWED_DIRS;
  else process.env.ASTRA_ALLOWED_DIRS = prevRoots;
  if (prevTmp === undefined) delete process.env.ASTRA_TMPDIR;
  else process.env.ASTRA_TMPDIR = prevTmp;
  fs.rmSync(dir, { recursive: true, force: true });
});

function parts(buf: Buffer): Map<string, string> {
  return new Map(new AdmZip(buf).getEntries().filter((e) => !e.isDirectory).map((e) => [e.entryName, sha(e.getData())]));
}

function copyOf(file: string, name: string): string {
  const target = path.join(dir, name);
  fs.copyFileSync(file, target);
  return target;
}

describe('CLI PDF export of a file with element-order errors', () => {
  it('exports from a verified copy: exit 0, a PDF, a note, the source file untouched, only document.xml changed', async () => {
    const source = copyOf(FIXTURE, 'order.docx');
    const before = sha(fs.readFileSync(source));
    const out = path.join(dir, 'out', 'order.pdf');
    const code = await runCli(['export', source, '--format', 'pdf', '-o', out]);
    expect(code).toBe(EXIT_OK);
    expect(fs.readFileSync(out).subarray(0, 5).toString()).toBe('%PDF-');
    expect(sha(fs.readFileSync(source))).toBe(before);
    const received = fs.readFileSync(captured);
    expect((await validateDocxBuffer(received)).ok).toBe(true);
    const a = parts(fs.readFileSync(source));
    const b = parts(received);
    const changed = [...new Set([...a.keys(), ...b.keys()])].filter((n) => a.get(n) !== b.get(n));
    expect(changed).toEqual(['word/document.xml']);
    expect(logged.join('\n')).toMatch(/note: the source has 1 element-order error\(s\) in word\/document\.xml/);
  });

  it('a mixed file (order error plus a real error) is refused: exit 1, no PDF, the backend never started', async () => {
    const zip = new AdmZip(fs.readFileSync(FIXTURE));
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('</w:body>', '<w:p><w:r><w:bogus/></w:r></w:p></w:body>'), 'utf8'));
    const source = path.join(dir, 'mixed.docx');
    fs.writeFileSync(source, zip.toBuffer());
    const out = path.join(dir, 'out', 'mixed.pdf');
    expect(await runCli(['export', source, '--format', 'pdf', '-o', out])).toBe(EXIT_FAIL);
    expect(fs.existsSync(out)).toBe(false);
    expect(fs.existsSync(path.join(dir, 'out'))).toBe(false);
    expect(fs.existsSync(marker)).toBe(false);
    expect(logged.join('\n')).toContain('astra normalize');
  });
});

describe('CLI PDF export of a valid file', () => {
  it('the backend receives the original bytes', async () => {
    const source = path.join(dir, 'good.docx');
    fs.copyFileSync(path.join(GALLERY, 'T2_edit_target.docx'), source);
    expect(await runCli(['export', source, '--format', 'pdf', '-o', path.join(dir, 'good.pdf')])).toBe(EXIT_OK);
    expect(sha(fs.readFileSync(captured))).toBe(sha(fs.readFileSync(source)));
    expect(logged.join('\n')).not.toMatch(/note: the source has/);
  });

  it('a package over the ratio ceiling is refused with E_LIMIT before the backend starts, valid or not', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    let seed = 99;
    const noise = Array.from({ length: 6000 }, () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) >> 8) & 0xff);
    zip.addFile('customXml/big.xml', Buffer.from('<r>' + Buffer.from(noise).toString('base64') + '</r>', 'utf8'));
    const source = path.join(dir, 'big.docx');
    fs.writeFileSync(source, zip.toBuffer());
    const prev = process.env.ASTRA_MAX_RATIO;
    process.env.ASTRA_MAX_RATIO = '1.05';
    try {
      expect(await runCli(['export', source, '--format', 'pdf', '-o', path.join(dir, 'big.pdf')])).toBe(EXIT_FAIL);
      expect(logged.join('\n')).toContain('E_LIMIT');
      expect(fs.existsSync(marker)).toBe(false);
    } finally {
      if (prev === undefined) delete process.env.ASTRA_MAX_RATIO;
      else process.env.ASTRA_MAX_RATIO = prev;
    }
  });
});

describe('MCP export_document pdf', () => {
  it('exports the order-error file (the session is serialized as a save would write it)', async () => {
    const source = copyOf(FIXTURE, 'mcp-order.docx');
    const store = createMcpSessionStore();
    const opened = (await handleOpenDocument(store, { path: source })) as { sessionId?: string };
    expect(opened.sessionId).toBeTruthy();
    const out = path.join(dir, 'mcp-order.pdf');
    const res = (await handleExportDocument(store, { sessionId: opened.sessionId!, format: 'pdf', outPath: out })) as { ok: boolean };
    expect(res.ok).toBe(true);
    expect(fs.readFileSync(out).subarray(0, 5).toString()).toBe('%PDF-');
    expect((await validateDocxBuffer(fs.readFileSync(captured))).ok).toBe(true);
  });

  it('contract: the PDF source equals what save_document writes for the same session, part by part', async () => {
    const source = copyOf(FIXTURE, 'mcp-contract.docx');
    const store = createMcpSessionStore();
    const opened = (await handleOpenDocument(store, { path: source })) as { sessionId?: string };
    const out = path.join(dir, 'mcp-contract.pdf');
    expect(((await handleExportDocument(store, { sessionId: opened.sessionId!, format: 'pdf', outPath: out })) as { ok: boolean }).ok).toBe(true);
    const exported = parts(fs.readFileSync(captured));
    const saved = path.join(dir, 'saved.docx');
    const res = (await handleSaveDocument(store, { sessionId: opened.sessionId!, outPath: saved })) as { ok?: boolean };
    expect(res.ok).not.toBe(false);
    const written = parts(fs.readFileSync(saved));
    expect([...exported.keys()].sort()).toEqual([...written.keys()].sort());
    for (const name of ['word/document.xml', 'word/styles.xml', '[Content_Types].xml']) {
      expect(exported.get(name)).toBe(written.get(name));
    }
  });

  it('refuses an invalid session with E_VALIDATION: no PDF, backend never started', async () => {
    const zip = new AdmZip(fs.readFileSync(path.join(GALLERY, 'T2_edit_target.docx')));
    const xml = zip.readAsText('word/document.xml');
    const at = xml.lastIndexOf('<w:sectPr');
    zip.updateFile('word/document.xml', Buffer.from(`${xml.slice(0, at)}<w:p><w:r><w:bogus/></w:r></w:p>${xml.slice(at)}`, 'utf8'));
    const bad = path.join(dir, 'bad.docx');
    fs.writeFileSync(bad, zip.toBuffer());
    const store = createMcpSessionStore();
    const opened = (await handleOpenDocument(store, { path: bad })) as { sessionId?: string };
    const out = path.join(dir, 'bad.pdf');
    const res = (await handleExportDocument(store, { sessionId: opened.sessionId!, format: 'pdf', outPath: out })) as { ok: boolean; code?: string };
    expect(res.ok).toBe(false);
    expect(res.code).toBe('E_VALIDATION');
    expect(fs.existsSync(out)).toBe(false);
    expect(fs.existsSync(marker)).toBe(false);
  });
});
