#!/usr/bin/env tsx
/**
 * Adds the six platform packages to the validator wrapper's optionalDependencies.
 *
 * The repository manifest deliberately does not list them: they are built from the
 * native binaries at release time and do not exist on npm until the first release,
 * so listing them would make `npm ci` fail ("Missing from lock file"). The publish
 * workflow runs this after the platform packages are published and before the
 * wrapper is, so the published wrapper carries them at its validatorBinary range.
 *
 *   npx tsx scripts/inject-validator-optional-deps.ts [path/to/package.json]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { VALIDATOR_TARGETS } from './validator-targets.js';
import { parseValidatorBinary } from './validator-binary.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const defaultManifest = path.resolve(here, '../packages/ooxml-validator/package.json');

/**
 * The optionalDependencies a published wrapper must carry: every platform package, at the wrapper's
 * binary range (`validatorBinary.range`, see validator-binary.ts). The range is not tied to the
 * wrapper's own version, so a fix to one platform's binary can be published as a new binary version
 * of that package alone and reach users without a new wrapper, and a platform whose rebuild failed
 * keeps the last binary published inside the range.
 */
export function platformOptionalDependencies(range: string): Record<string, string> {
  return Object.fromEntries(VALIDATOR_TARGETS.map((t) => [`@attentionform/ooxml-validator-${t.npmSuffix}`, range]));
}

/** Writes the platform packages into the manifest at `manifestPath` and returns the result. */
export function injectPlatformDependencies(manifestPath: string = defaultManifest): Record<string, unknown> {
  const pkg = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  if (typeof pkg.version !== 'string' || pkg.version === '') {
    throw new Error(`${manifestPath} has no version`);
  }
  const binary = parseValidatorBinary(pkg, manifestPath);
  pkg.optionalDependencies = { ...(pkg.optionalDependencies ?? {}), ...platformOptionalDependencies(binary.range) };
  fs.writeFileSync(manifestPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');
  return pkg;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const target = process.argv[2] ? path.resolve(process.argv[2]) : defaultManifest;
  const pkg = injectPlatformDependencies(target) as { optionalDependencies: Record<string, string>; validatorBinary: { range: string } };
  console.log(`[inject-validator-optional-deps] ${Object.keys(pkg.optionalDependencies).length} platform packages at ${pkg.validatorBinary.range} in ${target}`);
}
