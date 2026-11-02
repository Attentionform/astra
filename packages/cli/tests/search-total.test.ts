import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/**
 * `search --limit` cuts the list of hits, and the footer used to count only what was shown, so a list of 20 out of
 * 57 read as 20 matches in all. The footer now reads `N of M matches` when the list is cut, and `--json` has `total`.
 */
const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');
const run = (args: string[]) => {
  const r = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: r.status ?? 99, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
};

describe('search total', () => {
  let dir: string;
  let doc: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-search-total-'));
    const paragraphs = Array.from({ length: 7 }, (_, i) => `Hit number ${i + 1} holds the needle word.`);
    fs.writeFileSync(path.join(dir, 's.md'), `# Title\n\n${paragraphs.join('\n\n')}\n\nA closing paragraph without it.\n`);
    doc = path.join(dir, 's.docx');
    expect(run(['create', doc, '--from-md', path.join(dir, 's.md')]).status).toBe(0);
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  it('a limit below the number of hits shows N of M and says how to see more', () => {
    const r = run(['search', doc, 'needle', '--limit', '3']);
    expect(r.status).toBe(0);
    const lines = r.stdout.trim().split('\n');
    expect(lines).toHaveLength(4);
    expect(lines[3]).toBe('3 of 7 matches; raise --limit (1..100) to see more');
  });

  it('a limit that is not reached keeps the plain footer', () => {
    expect(run(['search', doc, 'needle']).stdout.trim().split('\n').pop()).toBe('7 matches');
    expect(run(['search', doc, 'needle', '--limit', '7']).stdout.trim().split('\n').pop()).toBe('7 matches');
    expect(run(['search', doc, 'closing paragraph']).stdout.trim().split('\n').pop()).toBe('1 match');
    expect(run(['search', doc, 'absent-word']).stdout.trim()).toBe('0 matches');
  });

  it('a limit of one shows 1 of M', () => {
    expect(run(['search', doc, 'needle', '--limit', '1']).stdout.trim().split('\n').pop()).toBe('1 of 7 matches; raise --limit (1..100) to see more');
  });

  it('--json carries total and truncated next to count', () => {
    const cut = JSON.parse(run(['--json', 'search', doc, 'needle', '--limit', '2']).stdout);
    expect(cut).toMatchObject({ ok: true, count: 2, total: 7, truncated: true });
    expect(cut.hits).toHaveLength(2);
    const whole = JSON.parse(run(['--json', 'search', doc, 'needle']).stdout);
    expect(whole).toMatchObject({ count: 7, total: 7, truncated: false });
    const none = JSON.parse(run(['--json', 'search', doc, 'absent-word']).stdout);
    expect(none).toMatchObject({ count: 0, total: 0, truncated: false });
  });

  it('the help describes the footer and the new fields', () => {
    const help = run(['search', '--help']).stdout.replace(/\s+/g, ' ');
    expect(help).toContain('the footer reads `N of M matches` and --json carries `total` (all matches) and `truncated`');
  });
});
