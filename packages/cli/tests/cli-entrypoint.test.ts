import { describe, it, expect, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const REPO_ROOT = path.resolve(__dirname, '..', '..', '..');
const CLI_JS = path.join(REPO_ROOT, 'packages', 'cli', 'dist', 'cli.js');
const MCP_JS = path.join(REPO_ROOT, 'packages', 'mcp-server', 'dist', 'index.js');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-entry-'));
afterAll(() => fs.rmSync(tmp, { recursive: true, force: true }));

function runCopy(source: string, relativeTarget: string, args: string[]) {
  const target = path.join(tmp, relativeTarget);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.copyFileSync(source, target);
  // The built files leave their packages external; let them find the ones in the repository.
  return spawnSync(process.execPath, [target, ...args], {
    encoding: 'utf8',
    env: { ...process.env, NODE_PATH: path.join(REPO_ROOT, 'node_modules') },
  });
}

/**
 * The npm shim on Windows starts a package as `node <install folder>\dist\cli.js`. That path is not
 * inside this repository and is not a file called `astra`, so a check that only looked at the path
 * made the CLI exit 0 and print nothing. The built files must run from any location and any name.
 */
describe('built entry points run from any path', () => {
  it('the CLI prints its version when node is given a copy of dist/cli.js elsewhere', () => {
    const res = runCopy(CLI_JS, path.join('node_modules', '@attentionform', 'astra', 'dist', 'cli.js'), ['--version']);
    expect(res.status).toBe(0);
    expect(res.stdout.trim()).toMatch(/\d+\.\d+\.\d+/);
  });

  it('the CLI prints its usage under a different file name', () => {
    const res = runCopy(CLI_JS, path.join('other', 'run-astra.js'), ['--help']);
    expect(res.status).toBe(0);
    expect(res.stdout).toContain('Usage: astra');
  });

  it('the MCP server prints its version under a different file name', () => {
    const res = runCopy(MCP_JS, path.join('other', 'server.js'), ['--version']);
    expect(res.status).toBe(0);
    expect(res.stdout.trim()).toMatch(/\d+\.\d+\.\d+/);
  });
});
