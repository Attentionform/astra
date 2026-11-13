import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/** An operation that the schema gate refuses names the closest action or key and points to `batch --schema`. */
const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');

function run(args: string[]): { status: number; stdout: string; stderr: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

describe('batch with an operation the schema refuses', () => {
  let dir: string;
  let doc: string;
  let target: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-unknown-action-'));
    doc = path.join(dir, 'a.docx');
    fs.writeFileSync(path.join(dir, 'a.md'), '# Title\n\nBody text.\n');
    expect(run(['create', doc, '--from-md', path.join(dir, 'a.md')]).status).toBe(0);
    target = run(['inspect', doc]).stdout.split('\n').find((l) => l.includes('Body text.'))!.split('|')[0].trim();
  });

  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  function batch(operation: Record<string, unknown>, json = false) {
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [operation] }));
    return run([...(json ? ['--json'] : []), 'batch', doc, ops, '--dry-run']);
  }

  it('"remove" suggests "delete" and points to the schema', () => {
    const r = batch({ action: 'remove', target });
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('Did you mean "delete"?');
    expect(r.stderr).toContain('Run `astra batch --schema` for an example of each action');
  });

  it('"setTable" suggests "table" and lists its keys and operations', () => {
    const r = batch({ action: 'setTable', target });
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('"setTable" is not an action. Did you mean "table"?');
    expect(r.stderr).toContain('Keys of "table":');
    expect(r.stderr).toContain('Operations of "table":');
  });

  it('"spaceAfter" on style suggests "spacingAfter" and lists the keys of style', () => {
    const r = batch({ action: 'style', target, spaceAfter: 120 });
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('"spaceAfter" is not a key of "style". Did you mean "spacingAfter"?');
    expect(r.stderr).toContain('Keys of "style":');
  });

  it('the --json error carries the same remedy', () => {
    const r = batch({ action: 'remove', target }, true);
    const text = r.stdout + r.stderr;
    expect(text).toContain('Did you mean \\"delete\\"?');
  });

  it('a valid operation is untouched', () => {
    expect(batch({ action: 'style', target, bold: true }).status).toBe(0);
  });
});
