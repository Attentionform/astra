import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { createMcpSessionStore } from '../../mcp-server/src/lib/session-store';
import { handleExportDocument } from '../../mcp-server/src/tools/export';

// Cross-surface parity probe (standing suite): the same footnoted document
// through CLI export and MCP export must agree. Compares HTML bodies modulo
// the documented envelope differences (MCP returns inline content while the
// CLI writes a file; each side derives <title> from its own input path, so
// the title tag is normalized before the diff). In-process via run(argv).
const dirs: string[] = [];

function tmpDir(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `${prefix}-`));
  dirs.push(dir);
  return dir;
}

async function runCli(args: string[]): Promise<{ status: number; stdout: string; stderr: string }> {
  vi.mocked(console.log).mockClear();
  vi.mocked(console.error).mockClear();
  const status = await run(args);
  return {
    status,
    stdout: vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n'),
    stderr: vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n'),
  };
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  for (const dir of dirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

const MANUSCRIPT = [
  '# Parity Probe',
  '',
  'Opening paragraph with a footnote[^fn] in running text.',
  '',
  '[^fn]: Vindication lineage parcel 7XQ.',
  '',
].join('\n');

function normalizeHtml(html: string): string {
  return html
    .replace(/\r\n|\r/g, '\n')
    .replace(/<title>.*<\/title>/, '<title>TITLE</title>')
    .trim();
}

describe('CLI/MCP export parity', () => {
  it('footnote bodies agree byte-for-byte between CLI and MCP HTML exports', async () => {
    const dir = tmpDir('export-parity');
    const md = path.join(dir, 'manuscript.md');
    fs.writeFileSync(md, MANUSCRIPT);
    const docx = path.join(dir, 'parity.docx');
    expect((await runCli(['import', md, '-o', docx])).status).toBe(0);

    const cliHtml = path.join(dir, 'cli.html');
    const cliRes = await runCli(['export', docx, '--format', 'html', '-o', cliHtml]);
    expect(cliRes.status).toBe(0);
    const cliText = fs.readFileSync(cliHtml, 'utf8');

    const store = createMcpSessionStore();
    const buf = fs.readFileSync(docx);
    const h = await store.openBuffer(buf, 'parity.docx');
    const mcp = await handleExportDocument(store, { sessionId: h.sessionId, format: 'html' });
    expect(mcp.ok).toBe(true);
    if (!mcp.ok) return;
    const mcpText = mcp.content ?? '';

    const body = 'Vindication lineage parcel 7XQ.';
    for (const [label, text] of [['CLI', cliText], ['MCP', mcpText]] as const) {
      expect(text, `${label} HTML carries the footnote body`).toContain(body);
      expect(text, `${label} HTML has no footnote placeholder`).not.toContain('see document');
    }
    expect(normalizeHtml(mcpText)).toBe(normalizeHtml(cliText));
  });
});
