#!/usr/bin/env tsx
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { VALIDATOR_TARGETS } from './validator-targets.js';
import { parseValidatorBinary } from './validator-binary.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const wrapperPkgPath = path.resolve(__dirname, '../packages/ooxml-validator/package.json');

export function checkOrBumpValidatorVersion(newVersion?: string): { version: string; ok: boolean } {
  if (!fs.existsSync(wrapperPkgPath)) {
    throw new Error(`Wrapper package.json not found: ${wrapperPkgPath}`);
  }

  const pkg = JSON.parse(fs.readFileSync(wrapperPkgPath, 'utf8'));

  if (newVersion) {
    pkg.version = newVersion;
    fs.writeFileSync(wrapperPkgPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');
    console.log(`[bump-validator-version] Updated validator version to ${newVersion}`);
  }

  const currentVersion = pkg.version;
  if (typeof currentVersion !== 'string' || currentVersion === '') {
    throw new Error('Wrapper package.json has no version');
  }
  // The six platform packages are not listed in the repository manifest. They are added at the
  // wrapper's binary range when publishing (scripts/inject-validator-optional-deps.ts) and carry the
  // binary version, which moves independently of the wrapper's own version (validator-binary.ts).
  const binary = parseValidatorBinary(pkg, wrapperPkgPath);

  console.log(`[bump-validator-version] Verified: wrapper version ${currentVersion}; the ${VALIDATOR_TARGETS.length} platform packages are at binary ${binary.version}, accepted range ${binary.range}`);
  return { version: currentVersion, ok: true };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const arg = process.argv[2];
  if (arg && arg !== '--check') {
    checkOrBumpValidatorVersion(arg);
  } else {
    checkOrBumpValidatorVersion();
  }
}
