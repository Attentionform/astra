/**
 * Licence consistency: the root LICENSE is the Business Source License 1.1,
 * every published package carries a byte-identical copy, every package.json
 * declares the SPDX id BUSL-1.1, and the NOTICE files point at the BSL.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'node:fs';
import * as path from 'node:path';

const ROOT = path.resolve(__dirname, '..', '..');
const PACKAGES = ['engine', 'cli', 'mcp-server', 'ooxml-validator'];
const read = (...p: string[]) => fs.readFileSync(path.join(ROOT, ...p), 'utf8');

describe('licence consistency', () => {
  const rootLicense = read('LICENSE');

  it('root LICENSE is the BSL 1.1 with the ASTra parameters', () => {
    expect(rootLicense.startsWith('Business Source License 1.1')).toBe(true);
    expect(rootLicense).toMatch(/^Licensor:\s+Attentionform/m);
    expect(rootLicense).toMatch(/^Additional Use Grant:/m);
    expect(rootLicense).toMatch(/^Change Date:/m);
    expect(rootLicense).toMatch(/^Change License:\s+Apache License, Version 2\.0/m);
  });

  it('keeps the standard BSL terms unmodified', () => {
    const terms = rootLicense.slice(rootLicense.indexOf('\nTerms'));
    for (const clause of [
      'The Licensor hereby grants you the right to copy, modify, create derivative',
      'Not to modify this License in any other way.',
      'is not an Open',
    ]) {
      expect(terms).toContain(clause);
    }
  });

  it.each(PACKAGES)('package %s has a byte-identical LICENSE and the BUSL-1.1 id', (name) => {
    expect(read('packages', name, 'LICENSE')).toBe(rootLicense);
    const pkg = JSON.parse(read('packages', name, 'package.json'));
    expect(pkg.license).toBe('BUSL-1.1');
    expect(pkg.files).toContain('LICENSE');
  });

  it('root package.json declares BUSL-1.1', () => {
    expect(JSON.parse(read('package.json')).license).toBe('BUSL-1.1');
  });

  it.each(['.', ...PACKAGES.map((n) => path.join('packages', n))])('NOTICE in %s names the BSL', (dir) => {
    const notice = read(dir, 'NOTICE');
    expect(notice).toContain('Business Source License 1.1');
    expect(notice).not.toMatch(/licensed under the Apache License/i);
  });

  it('the documents do not call ASTra open source', () => {
    for (const f of ['README.md', 'LICENSING.md', 'CONTRIBUTING.md', path.join('docs', 'contributing.md')]) {
      const prose = read(f).replace(/`[^`]*`/g, '');
      expect(prose, f).not.toMatch(/\bASTra is an? open[- ]source\b/i);
    }
  });
});
