import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/** The history help names `.astra-tombstones.json`, and what it says about the file is what happens. */
const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');
const run = (args: string[], cwd?: string) => {
  const r = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8', cwd });
  return { status: r.status ?? 99, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
};

describe('history help and the tombstone file', () => {
  let dir: string;
  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-tombstone-'));
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  it('the help names the file next to the .astra/ folder, and says what to do with both', () => {
    const help = run(['history', '--help']).stdout.replace(/\s+/g, ' ');
    expect(help).toContain('.astra-tombstones.json');
    expect(help).toContain('sits beside the .astra/ folder');
    expect(help).toContain('lists the documents that have had a history');
    expect(help).toContain('add both to .gitignore');
  });

  it('history --init writes the file beside .astra/, and deleting .astra/ gives the warning the help describes', () => {
    const doc = path.join(dir, 'h.docx');
    expect(run(['create', doc]).status).toBe(0);
    expect(run(['history', doc, '--init']).status).toBe(0);
    expect(fs.existsSync(path.join(dir, '.astra'))).toBe(true);
    expect(fs.existsSync(path.join(dir, '.astra-tombstones.json'))).toBe(true);
    const registry = JSON.parse(fs.readFileSync(path.join(dir, '.astra-tombstones.json'), 'utf8'));
    expect(JSON.stringify(registry)).toContain('h.docx');

    fs.rmSync(path.join(dir, '.astra'), { recursive: true, force: true });
    const again = run(['history', doc, '--init']);
    expect(again.status).toBe(0);
    expect(again.stderr).toContain('possible history removal');
  });
});
