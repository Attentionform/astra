import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createBlankDocxBuffer, DocxParser, OperationEngine } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function writeBlankDocx(prefix = 'astra-stale-anchor'): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = tmpPath(prefix);
  fs.writeFileSync(file, buf);
  return file;
}

function writeOps(payload: unknown): string {
  const file = tmpPath('astra-stale-ops', '.json');
  fs.writeFileSync(file, JSON.stringify(payload));
  return file;
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
});

describe('stale anchor remedy (session-scoped, no disk-stable promise)', () => {
  it('engine batch stale anchor hint names search + anchor:end', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'stale_anchor_xyz', position: 'after', text: 'nope' } as any,
      ],
      sequential: true,
    });
    expect(res.success).toBe(false);
    const combined = `${res.errors[0].error} ${res.errors[0].hint ?? ''}`;
    expect(combined).toContain('stale_anchor_xyz');
    expect((res.errors[0].hint ?? '').toLowerCase()).toContain('search');
    expect(res.errors[0].hint ?? '').toContain('anchor:end');
    expect(res.errors[0].hint ?? '').toContain('session-scoped');
    expect(res.errors[0].hint ?? '').toContain('inspect_document');
    expect(res.errors[0].hint ?? '').toContain('search_content');
    expect(res.errors[0].hint ?? '').not.toContain('get_document_toon');
  });

  it('CLI batch stale anchor error names search + anchor:end', async () => {
    const file = writeBlankDocx('astra-stale-cli');
    const ops = writeOps({
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'stale_anchor_cli_xyz', position: 'after', text: 'nope' }],
      sequential: true,
    });
    vi.mocked(console.error).mockClear();
    await expect(run(['batch', file, ops])).resolves.toBe(1);
    const errText = vi.mocked(console.error).mock.calls.map((c) => String(c[0])).join('\n');
    expect(errText).toContain('stale_anchor_cli_xyz');
    expect(errText.toLowerCase()).toContain('search');
    expect(errText).toContain('anchor:end');
  });

  it('batch --help documents the refresh workflow', async () => {
    vi.mocked(console.log).mockClear();
    await expect(run(['batch', '--help'])).resolves.toBe(0);
    const out = vi.mocked(console.log).mock.calls.map((c) => String(c[0])).join('\n');
    expect(out.toLowerCase()).toContain('session-scoped');
    expect(out).toContain('astra inspect');
    expect(out.toLowerCase()).toContain('search');
    expect(out).toContain('anchor:end');
    expect(out.toLowerCase()).toContain('stale');
  });
});
