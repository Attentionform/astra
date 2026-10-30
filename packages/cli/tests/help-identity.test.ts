import { describe, it, expect } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';

const REPO_ROOT = path.resolve(__dirname, '..', '..', '..');
const CLI_JS = path.join(REPO_ROOT, 'packages', 'cli', 'dist', 'cli.js');
const GUIDE = path.join(REPO_ROOT, 'docs', 'cli-guide.md');

function runCli(args: string[]): { status: number; stdout: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '' };
}

function extractFencedBlocks(md: string): string[] {
  const out: string[] = [];
  const re = /```text\n([\s\S]*?)\n```/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(md)) !== null) out.push(m[1]);
  return out;
}

describe('help identity (both help sites + guide quote byte-identical)', () => {
  it('astra --help equals astra help and matches guide §2 verbatim', () => {
    const top = runCli(['--help']);
    expect(top.status).toBe(0);
    const viaHelp = runCli(['help']);
    expect(viaHelp.status).toBe(0);
    expect(viaHelp.stdout).toBe(top.stdout);

    const md = fs.readFileSync(GUIDE, 'utf8');
    const blocks = extractFencedBlocks(md);
    expect(blocks.length).toBeGreaterThanOrEqual(2);
    // §2 is the first ```text block in the guide.
    const guideTop = blocks[0];
    // Spawn adds trailing newline; guide fence strips it — compare trimmed-end.
    expect(top.stdout.replace(/\r\n/g, '\n').trimEnd()).toBe(guideTop.replace(/\r\n/g, '\n').trimEnd());
    // Byte-identical after normalizing the single trailing newline.
    expect(top.stdout.replace(/\r\n/g, '\n').trimEnd() + '\n').toBe(guideTop.replace(/\r\n/g, '\n').trimEnd() + '\n');
  });
});
