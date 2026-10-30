import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';

describe('third-party-notices sync and completeness', () => {
  const rootNoticesPath = path.resolve(__dirname, '../../THIRD_PARTY_NOTICES.md');
  const rootNoticesContent = fs.readFileSync(rootNoticesPath, 'utf8');

  it('root THIRD_PARTY_NOTICES.md contains entries for xarsh, Open XML SDK, and .NET runtime', () => {
    expect(rootNoticesContent).toContain('xarsh/ooxml-validator');
    expect(rootNoticesContent).toContain('Open XML SDK (DocumentFormat.OpenXml)');
    expect(rootNoticesContent).toContain('.NET runtime');
    expect(rootNoticesContent).toContain('MIT License (xarsh/ooxml-validator)');
    expect(rootNoticesContent).toContain('MIT License (Open XML SDK and .NET runtime)');
  });

  it('every package listing THIRD_PARTY_NOTICES.md in files has a byte-identical copy', () => {
    const packagesDir = path.resolve(__dirname, '../../packages');
    const packages = fs.readdirSync(packagesDir);

    for (const pkg of packages) {
      const pkgJsonPath = path.join(packagesDir, pkg, 'package.json');
      if (!fs.existsSync(pkgJsonPath)) continue;

      const pkgJson = JSON.parse(fs.readFileSync(pkgJsonPath, 'utf8'));
      const files: string[] = pkgJson.files || [];

      if (files.includes('THIRD_PARTY_NOTICES.md')) {
        const pkgNoticesPath = path.join(packagesDir, pkg, 'THIRD_PARTY_NOTICES.md');
        expect(fs.existsSync(pkgNoticesPath), `Missing THIRD_PARTY_NOTICES.md in packages/${pkg}`).toBe(true);

        const pkgNoticesContent = fs.readFileSync(pkgNoticesPath, 'utf8');
        expect(pkgNoticesContent, `packages/${pkg}/THIRD_PARTY_NOTICES.md must match root file`).toBe(rootNoticesContent);
      }
    }
  });
});
