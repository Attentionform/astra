/**
 * The version of the native validator binary, kept apart from the version of the wrapper package.
 *
 * The six platform packages (`@attentionform/ooxml-validator-<platform>`) carry the BINARY version. It
 * changes only when the native validator changes. The wrapper (`@attentionform/ooxml-validator`)
 * declares in its package.json:
 *
 *   "validatorBinary": { "version": "0.1.0", "range": ">=0.1.0 <0.2.0" }
 *
 * `version` is what the next platform release publishes. `range` is every binary the wrapper works
 * with, and it is what the wrapper's optionalDependencies ask for. Because the range is not tied to
 * the wrapper's own version, a platform whose rebuild failed keeps the last binary published inside
 * the range. Move the range up (and release all six platforms) only when the wrapper and the binary
 * stop being able to talk to each other.
 *
 * Only the range form `>=X.Y.Z <A.B.C` is allowed, so it means the same thing here as in npm.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
export const WRAPPER_MANIFEST = path.resolve(here, '../packages/ooxml-validator/package.json');

export interface ValidatorBinary {
  version: string;
  range: string;
}

const VERSION = /^(\d+)\.(\d+)\.(\d+)$/;
const RANGE = /^>=(\d+\.\d+\.\d+) <(\d+\.\d+\.\d+)$/;

/** Negative, zero or positive like a comparator. Both must be plain X.Y.Z versions. */
export function compareVersions(a: string, b: string): number {
  const pa = VERSION.exec(a);
  const pb = VERSION.exec(b);
  if (!pa || !pb) throw new Error(`Not a plain X.Y.Z version: ${!pa ? a : b}`);
  for (let i = 1; i <= 3; i++) {
    const diff = Number(pa[i]) - Number(pb[i]);
    if (diff !== 0) return diff;
  }
  return 0;
}

/** True when `version` is a plain X.Y.Z inside a `>=X.Y.Z <A.B.C` range; false for anything else (pre-releases included). */
export function inBinaryRange(version: string, range: string): boolean {
  const m = RANGE.exec(range);
  if (!m || !VERSION.test(version)) return false;
  return compareVersions(version, m[1]) >= 0 && compareVersions(version, m[2]) < 0;
}

export function parseValidatorBinary(manifest: unknown, where = 'package.json'): ValidatorBinary {
  const field = (manifest as { validatorBinary?: { version?: unknown; range?: unknown } } | null)?.validatorBinary;
  if (!field || typeof field.version !== 'string' || typeof field.range !== 'string') {
    throw new Error(`${where} needs "validatorBinary": { "version": "X.Y.Z", "range": ">=X.Y.Z <A.B.C" }`);
  }
  if (!VERSION.test(field.version)) throw new Error(`${where}: validatorBinary.version must be a plain X.Y.Z version, got "${field.version}"`);
  if (!RANGE.test(field.range)) throw new Error(`${where}: validatorBinary.range must look like ">=0.1.0 <0.2.0", got "${field.range}"`);
  if (!inBinaryRange(field.version, field.range)) {
    throw new Error(`${where}: validatorBinary.version ${field.version} is outside its own range ${field.range}`);
  }
  return { version: field.version, range: field.range };
}

export function readValidatorBinary(manifestPath: string = WRAPPER_MANIFEST): ValidatorBinary {
  return parseValidatorBinary(JSON.parse(fs.readFileSync(manifestPath, 'utf8')), manifestPath);
}
