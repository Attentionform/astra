import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { RUNTIME_NOTICE_FILES, buildValidatorPackages } from '../../scripts/build-validator-packages';
import { VALIDATOR_TARGETS } from '../../scripts/validator-targets';

const ROOT = path.resolve(__dirname, '..', '..');

describe('validator platform packages', () => {
  let work: string;
  let binaries: string;
  let out: string;

  beforeEach(() => {
    work = fs.mkdtempSync(path.join(os.tmpdir(), 'validator-pkgs-'));
    binaries = path.join(work, 'binaries');
    out = path.join(work, 'out');
    for (const t of VALIDATOR_TARGETS) {
      const dir = path.join(binaries, t.rid);
      fs.mkdirSync(dir, { recursive: true });
      fs.writeFileSync(path.join(dir, t.binaryName), `fake ${t.rid} binary`);
      for (const n of RUNTIME_NOTICE_FILES) fs.writeFileSync(path.join(dir, n), `${n} for ${t.rid}`);
    }
  });

  afterEach(() => fs.rmSync(work, { recursive: true, force: true }));

  it('builds six packages that each ship the licence, notices and the .NET runtime notices', () => {
    const { builtTargets } = buildValidatorPackages({ binariesDir: binaries, outDir: out });
    expect(builtTargets).toHaveLength(6);
    const rootLicense = fs.readFileSync(path.join(ROOT, 'LICENSE'), 'utf8');
    for (const t of VALIDATOR_TARGETS) {
      const dir = path.join(out, t.npmSuffix);
      const pkg = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
      expect(pkg.license).toBe('BUSL-1.1');
      expect(fs.readFileSync(path.join(dir, 'LICENSE'), 'utf8')).toBe(rootLicense);
      for (const n of ['NOTICE', 'THIRD_PARTY_NOTICES.md', ...RUNTIME_NOTICE_FILES]) {
        expect(pkg.files, `${t.rid} files`).toContain(n);
        expect(fs.existsSync(path.join(dir, n)), `${t.rid} ${n}`).toBe(true);
      }
      expect(fs.readFileSync(path.join(dir, 'THIRD-PARTY-NOTICES.TXT'), 'utf8')).toContain(t.rid);
    }
  });

  it('gives every platform package the binary version, not the wrapper version', () => {
    buildValidatorPackages({ binariesDir: binaries, outDir: out });
    const wrapper = JSON.parse(fs.readFileSync(path.join(ROOT, 'packages/ooxml-validator/package.json'), 'utf8'));
    for (const t of VALIDATOR_TARGETS) {
      const pkg = JSON.parse(fs.readFileSync(path.join(out, t.npmSuffix, 'package.json'), 'utf8'));
      expect(pkg.version).toBe(wrapper.validatorBinary.version);
    }
  });

  it.each(RUNTIME_NOTICE_FILES)('refuses to build a package without %s', (missing) => {
    const target = VALIDATOR_TARGETS[0];
    fs.rmSync(path.join(binaries, target.rid, missing));
    expect(() => buildValidatorPackages({ binariesDir: binaries, outDir: out })).toThrow(missing);
  });

  it('the csproj copies the runtime notices on publish', () => {
    const csproj = fs.readFileSync(path.join(ROOT, 'packages/ooxml-validator/native/ooxml-validator.csproj'), 'utf8');
    expect(csproj).toContain('Name="CopyRuntimeNotices"');
    for (const n of ['THIRD-PARTY-NOTICES.TXT', 'DOTNET-LICENSE.TXT']) expect(csproj).toContain(n);
  });
});
