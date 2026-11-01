/**
 * A file that already has validator errors is written when the edit adds none, and the output says
 * how many errors were there before. --strict refuses every error. The field names match the MCP
 * save_document response for the same file.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../mcp-server/src/lib/session-store';
import { LEAN6_HANDLERS } from '../../mcp-server/src/lib/mcp-tool-adapter';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

/** Blank document with one style whose w:type is not allowed: a schema error nothing in the engine repairs. */
function writeDocx(prefix: string, invalid: boolean): string {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  if (invalid) {
    const styles = zip.getEntry('word/styles.xml')?.getData().toString('utf8') ?? '';
    const bad = '<w:style w:type="bogus" w:styleId="Bad1"><w:name w:val="Bad1"/></w:style>';
    zip.updateFile('word/styles.xml', Buffer.from(styles.replace('</w:styles>', `${bad}</w:styles>`), 'utf8'));
  }
  const file = tmpPath(prefix);
  fs.writeFileSync(file, zip.toBuffer());
  return file;
}

function writeOps(): string {
  const file = tmpPath('astra-baseline-ops', '.json');
  fs.writeFileSync(file, JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'hello' }] }));
  return file;
}

const stdoutText = (): string => vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
const stderrText = (): string => vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpFiles.length > 0) {
    try {
      fs.unlinkSync(tmpFiles.pop() as string);
    } catch {
      /* best-effort */
    }
  }
});

describe('batch on a file that already has validator errors', () => {
  it('writes the file and reports the errors that were already there', async () => {
    const file = writeDocx('astra-baseline-tolerated', true);
    await expect(run(['--json', 'batch', file, writeOps()])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText().split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(payload.ok).toBe(true);
    expect(payload.validation).toMatchObject({ tolerated: true, introducedErrors: 0 });
    expect(payload.validation.preExistingErrors).toBeGreaterThanOrEqual(1);
    expect(stderrText()).toContain('NOTE: saved with');
    expect(stderrText()).toContain('0 introduced');
  });

  it('refuses the same edit with --strict, as before', async () => {
    const file = writeDocx('astra-baseline-strict', true);
    const before = fs.readFileSync(file);
    await expect(run(['batch', file, writeOps(), '--strict'])).resolves.toBe(1);
    expect(stderrText()).toContain('E_VALIDATION');
    expect(fs.readFileSync(file).equals(before)).toBe(true);
  });

  it('adds no validation field for a valid file', async () => {
    const file = writeDocx('astra-baseline-valid', false);
    await expect(run(['--json', 'batch', file, writeOps()])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText().split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(payload.validation).toBeUndefined();
  });

  it('carries the same field names and numbers as the MCP save_document response', async () => {
    const cliFile = writeDocx('astra-baseline-parity-cli', true);
    await run(['--json', 'batch', cliFile, writeOps()]);
    const cli = JSON.parse(stdoutText().split('\n').filter((l) => l.startsWith('{')).pop() as string).validation;

    const store = createMcpSessionStore();
    const opened = await LEAN6_HANDLERS.open_document(store, { path: writeDocx('astra-baseline-parity-mcp', true) });
    if (!opened.ok) throw new Error('open failed');
    await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'hello' }],
    } as never);
    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId: opened.sessionId });
    if (!saved.ok) throw new Error('save failed');
    const mcp = saved.validation as Record<string, unknown>;

    expect(cli.tolerated).toBe(mcp.tolerated);
    expect(cli.preExistingErrors).toBe(mcp.preExistingErrors);
    expect(cli.introducedErrors).toBe(mcp.introducedErrors);
  });
});

describe('normalize on a file that already has validator errors', () => {
  it('writes the file and reports the errors that were already there', async () => {
    const file = writeDocx('astra-baseline-normalize', true);
    await expect(run(['--json', 'normalize', file])).resolves.toBe(0);
    const payload = JSON.parse(stdoutText().split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(payload.validation).toMatchObject({ tolerated: true, introducedErrors: 0 });
  });

  it('refuses with --strict', async () => {
    const file = writeDocx('astra-baseline-normalize-strict', true);
    await expect(run(['normalize', file, '--strict'])).resolves.toBe(1);
  });
});
