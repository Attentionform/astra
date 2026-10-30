import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip } from '../../packages/engine/src';
import { setPdfBackendBinaryForTesting, resetPdfBackendCache } from '../../packages/engine/src/export/pdf';
import { run as runCli, EXIT_FAIL } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleOpenDocument } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { handleExportDocument } from '../../packages/mcp-server/src/tools/export';

const GALLERY = path.resolve(__dirname, '../../test-gallery');

let dir: string;
let prevTmp: string | undefined;
let marker: string;
let prevRoots: string | undefined;

function invalidCopy(): string {
  const zip = new AdmZip(fs.readFileSync(path.join(GALLERY, 'T2_edit_target.docx')));
  const xml = zip.readAsText('word/document.xml');
  const at = xml.lastIndexOf('<w:sectPr');
  zip.updateFile('word/document.xml', Buffer.from(`${xml.slice(0, at)}<w:p><w:r><w:bogus/></w:r></w:p>${xml.slice(at)}`, 'utf8'));
  const file = path.join(dir, 'bad.docx');
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-pdf-val-'));
  prevTmp = process.env.ASTRA_TMPDIR;
  process.env.ASTRA_TMPDIR = path.join(dir, 'tmp');
  marker = path.join(dir, 'backend-ran');
  const fake = path.join(dir, 'fake-office');
  fs.writeFileSync(fake, `#!/bin/sh\ntouch "${marker}"\nexit 1\n`, { mode: 0o755 });
  setPdfBackendBinaryForTesting(fake);
  prevRoots = process.env.ASTRA_ALLOWED_DIRS;
  process.env.ASTRA_ALLOWED_DIRS = dir;
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
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

describe('PDF export validates its source before rendering', () => {
  it('CLI refuses an invalid docx: exit 1, no PDF, no directory, backend never started', async () => {
    const bad = invalidCopy();
    const outDir = path.join(dir, 'made-by-export');
    const code = await runCli(['export', bad, '--format', 'pdf', '-o', path.join(outDir, 'bad.pdf')]);
    expect(code).toBe(EXIT_FAIL);
    expect(fs.existsSync(outDir)).toBe(false);
    expect(fs.existsSync(marker)).toBe(false);
  });

  it('MCP refuses an invalid session with E_VALIDATION: no PDF, backend never started', async () => {
    const bad = invalidCopy();
    const store = createMcpSessionStore();
    const opened = await handleOpenDocument(store, { path: bad });
    const sessionId = (opened as { sessionId?: string }).sessionId;
    expect(sessionId).toBeTruthy();
    const out = path.join(dir, 'bad.pdf');
    const res = (await handleExportDocument(store, { sessionId: sessionId!, format: 'pdf', outPath: out })) as {
      ok: boolean;
      code?: string;
    };
    expect(res.ok).toBe(false);
    expect(res.code).toBe('E_VALIDATION');
    expect(fs.existsSync(out)).toBe(false);
    expect(fs.existsSync(marker)).toBe(false);
  });

  it('a valid docx still reaches the backend (the validation does not block it)', async () => {
    const good = path.join(dir, 'good.docx');
    fs.copyFileSync(path.join(GALLERY, 'T2_edit_target.docx'), good);
    await runCli(['export', good, '--format', 'pdf', '-o', path.join(dir, 'good.pdf')]);
    expect(fs.existsSync(marker)).toBe(true);
  });
});
