import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, it, expect } from 'vitest';
import { VALIDATOR_TARGETS } from '../../scripts/validator-targets.js';
import { injectPlatformDependencies } from '../../scripts/inject-validator-optional-deps.js';
import { inBinaryRange, parseValidatorBinary, readValidatorBinary } from '../../scripts/validator-binary.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const wrapperPkgPath = path.resolve(__dirname, '../../packages/ooxml-validator/package.json');

describe('Validator version sync (Milestone M5)', () => {
  it('the repository manifest lists no platform packages (they do not exist on npm until release)', () => {
    const pkg = JSON.parse(fs.readFileSync(wrapperPkgPath, 'utf8'));
    expect(pkg.version).toBeTruthy();
    expect(pkg.optionalDependencies).toBeUndefined();
  });

  it('the wrapper declares the binary it ships and the range of binaries it accepts', () => {
    const binary = readValidatorBinary(wrapperPkgPath);
    expect(binary.version).toBe('0.1.0');
    expect(inBinaryRange(binary.version, binary.range)).toBe(true);
  });

  it('publish-time injection adds every platform package at the wrapper binary range, not at its own version', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-inject-'));
    try {
      const manifest = path.join(dir, 'package.json');
      fs.copyFileSync(wrapperPkgPath, manifest);
      const { range } = readValidatorBinary(manifest);
      const out = injectPlatformDependencies(manifest) as { optionalDependencies: Record<string, string> };
      expect(Object.keys(out.optionalDependencies)).toHaveLength(VALIDATOR_TARGETS.length);
      for (const target of VALIDATOR_TARGETS) {
        expect(out.optionalDependencies[`@attentionform/ooxml-validator-${target.npmSuffix}`]).toBe(range);
      }
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it('consuming packages declare semver range matching wrapper version', () => {
    const pkg = JSON.parse(fs.readFileSync(wrapperPkgPath, 'utf8'));
    const expectedRange = `^${pkg.version}`;
    const consumers = ['engine', 'cli', 'mcp-server'];

    for (const name of consumers) {
      const consumerPkgPath = path.resolve(__dirname, `../../packages/${name}/package.json`);
      expect(fs.existsSync(consumerPkgPath)).toBe(true);
      const consumerPkg = JSON.parse(fs.readFileSync(consumerPkgPath, 'utf8'));
      const dep = consumerPkg.dependencies?.['@attentionform/ooxml-validator'];
      expect(dep, `packages/${name} dependency range`).toBe(expectedRange);
    }
  });

  it('a wrapper release does not move the platform versions: the range still accepts the binary after a patch bump', () => {
    const { version, range } = readValidatorBinary(wrapperPkgPath);
    const wrapperAfterPatch = { version: '0.1.5', validatorBinary: { version, range } };
    expect(parseValidatorBinary(wrapperAfterPatch).version).toBe(version);
    // the old scheme asked for ^<wrapper version>: it would have refused the binary that is on npm
    expect(inBinaryRange(version, `>=${'0.1.5'} <0.2.0`)).toBe(false);
  });

  it('refuses a binary version that is outside its own range, and ranges npm would read differently', () => {
    expect(() => parseValidatorBinary({ validatorBinary: { version: '0.3.0', range: '>=0.1.0 <0.2.0' } })).toThrow(/outside its own range/);
    expect(() => parseValidatorBinary({ validatorBinary: { version: '0.1.0', range: '^0.1.0' } })).toThrow(/must look like/);
    expect(() => parseValidatorBinary({ validatorBinary: { version: '0.1.0-beta.1', range: '>=0.1.0 <0.2.0' } })).toThrow(/plain X.Y.Z/);
    expect(() => parseValidatorBinary({})).toThrow(/needs "validatorBinary"/);
  });

  it('every published package asks for the same Node floor', () => {
    const floors = ['ooxml-validator', 'engine', 'cli', 'mcp-server'].map(
      (n) => JSON.parse(fs.readFileSync(path.resolve(__dirname, `../../packages/${n}/package.json`), 'utf8')).engines.node,
    );
    expect(new Set(floors)).toEqual(new Set(['>=22.19.0']));
  });
});
