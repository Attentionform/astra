/**
 * What npm shows and checks about each published package. A release from a public repository is
 * published with provenance, and npm refuses the publish when `repository.url` is missing or does not
 * name the repository the workflow runs in, so every package must carry it.
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, it, expect } from 'vitest';

const ROOT = path.resolve(__dirname, '../..');
const REPO = 'https://github.com/Attentionform/astra';
const PACKAGES = ['ooxml-validator', 'engine', 'cli', 'mcp-server'];

describe.each(PACKAGES)('packages/%s manifest', (dir) => {
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages', dir, 'package.json'), 'utf8'));

  it('names this repository and its own folder, as provenance requires', () => {
    expect(pkg.repository.type).toBe('git');
    expect(pkg.repository.url.replace(/^git\+/, '').replace(/\.git$/, '')).toBe(REPO);
    expect(pkg.repository.directory).toBe(`packages/${dir}`);
    expect(fs.existsSync(path.join(ROOT, pkg.repository.directory, 'package.json'))).toBe(true);
  });

  it('has the links, author, keywords and public access an npm page needs', () => {
    expect(pkg.homepage).toMatch(/^https:\/\//);
    expect(pkg.bugs.url).toBe(`${REPO}/issues`);
    expect(pkg.author.name).toBe('Attentionform');
    expect(pkg.keywords.length).toBeGreaterThan(0);
    expect(pkg.publishConfig).toEqual({ access: 'public' });
    expect(pkg.license).toBe('BUSL-1.1');
  });
});
