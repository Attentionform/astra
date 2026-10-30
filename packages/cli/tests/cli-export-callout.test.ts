import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';

// Persisted-callout export pins (CLI surface — the path users hit): every
// CLI export reopens from disk, so the parser rebuild (metadata.callout
// reconstructed from the saved table structure) is what turns a reopened
// callout back into a div. Plain data tables must stay tables.
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

async function firstNodeId(docx: string): Promise<string> {
  const inspected = await runCli(['inspect', docx]);
  expect(inspected.status).toBe(0);
  const line = inspected.stdout.split('\n').find((l) => /^[0-9A-F]{8} /.test(l));
  expect(line).toBeTruthy();
  return (line as string).split(' ')[0];
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

describe('CLI persisted callout export', () => {
  it('reopened callout exports as div.callout, never a data table', async () => {
    const dir = tmpDir('cli-callout-div');
    const docx = path.join(dir, 'memo.docx');
    expect((await runCli(['create', docx, '--template', 'blank'])).status).toBe(0);
    const anchor = await firstNodeId(docx);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(
      ops,
      JSON.stringify({
        operations: [
          { action: 'callout', anchor, position: 'after', title: 'North Gate', text: 'Keep the north gate shut.', theme: 'amber' },
        ],
      })
    );
    expect((await runCli(['batch', docx, ops])).status).toBe(0);
    const html = path.join(dir, 'memo.html');
    const exported = await runCli(['export', docx, '--format', 'html', '-o', html]);
    expect(exported.status).toBe(0);
    const body = fs.readFileSync(html, 'utf8');
    expect(body).toContain('<div class="callout"><h4>North Gate</h4><p>Keep the north gate shut.</p></div>');
    expect(body).not.toContain('<table');
  });

  it('reopened plain data tables still export as tables (no over-eager div)', async () => {
    const dir = tmpDir('cli-callout-negative');
    const docx = path.join(dir, 'data.docx');
    expect((await runCli(['create', docx, '--template', 'blank'])).status).toBe(0);
    const anchor = await firstNodeId(docx);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(
      ops,
      JSON.stringify({
        operations: [
          { action: 'insert', type: 'table', anchor, position: 'after', headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']] },
          { action: 'insert', type: 'table', anchor, position: 'after', headers: ['Solo'], rows: [] },
        ],
      })
    );
    expect((await runCli(['batch', docx, ops])).status).toBe(0);
    const html = path.join(dir, 'data.html');
    const exported = await runCli(['export', docx, '--format', 'html', '-o', html]);
    expect(exported.status).toBe(0);
    const body = fs.readFileSync(html, 'utf8');
    expect(body.match(/<table/g)?.length).toBe(2);
    expect(body).not.toContain('<div class="callout">');
  });
});
