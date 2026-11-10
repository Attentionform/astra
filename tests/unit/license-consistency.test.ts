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

  it('the grant carries the 0.4.0 limits and defines the three free-use paths', () => {
    const grant = rootLicense.slice(rootLicense.indexOf('Additional Use Grant:'), rootLicense.indexOf('Change Date:')).replace(/\s+/g, ' ');
    expect(grant).toContain('"Individual Use" means use by a natural person who works alone');
    expect(grant).toContain('"Internal Use" means use of any part of the Licensed Work');
    expect(grant).toContain('"Small Organization Use"');
    expect(grant).toContain('no more than 10 people working for it');
    expect(grant).toContain('"Small Product Use"');
    // a person who works alone is free for their own work, but a product of their own is judged by its revenue
    expect(grant).toContain('Individual Use does not include making the Licensed Work part of a product or service of Your own that You offer to others, whether or not You work alone');
    expect(grant).toContain('Individual Use never includes a Competing Offering');
    expect(grant.match(/USD 250,000/g)?.length).toBeGreaterThanOrEqual(2);
    expect(grant).not.toMatch(/USD 1,000,000|Small Business Use/);
    expect(grant).toContain("an employee who uses the Licensed Work for the employer's work is making the employer's use");
  });

  it('LICENSE and LICENSING.md carry no prices and no plan names; they live on the pricing page', () => {
    for (const f of ['LICENSE', 'LICENSING.md']) {
      const text = read(f);
      expect(text, f).not.toMatch(/USD\s?[0-9][0-9,]*\s*(per|a|\/)\s*(year|month)/i);
      expect(text, f).not.toMatch(/\b(Team Plus|Business Plus|Scale)\b/);
    }
    expect(read('LICENSING.md')).toContain('https://attentionform.com/pricing');
    expect(rootLicense).toContain('https://attentionform.com/pricing');
  });

  it('the documents do not call ASTra open source', () => {
    for (const f of ['README.md', 'LICENSING.md', 'CONTRIBUTING.md', path.join('docs', 'contributing.md')]) {
      const prose = read(f).replace(/`[^`]*`/g, '');
      expect(prose, f).not.toMatch(/\bASTra is an? open[- ]source\b/i);
    }
  });
});
