#!/usr/bin/env tsx
/**
 * Asks npm whether every validator platform has a binary the wrapper can use, before the wrapper is
 * published.
 *
 * The wrapper asks for each platform package at its binary range (validator-binary.ts), not at its own
 * version. So a platform whose newest binary failed to publish is fine as long as an older binary
 * inside the range is on npm: users on that platform keep it. A platform with NO binary in the range
 * would leave its users with a "Missing optional dependency" error, and that blocks the wrapper.
 *
 * --require-all is stricter: every platform must have the CURRENT binary version (the first release,
 * or a change to how the wrapper and the binary talk to each other).
 *
 *   npx tsx scripts/validator-platform-status.ts [--require-all]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { publishedVersions } from './npm-publish-if-new.js';
import { VALIDATOR_TARGETS } from './validator-targets.js';
import { compareVersions, inBinaryRange, readValidatorBinary, type ValidatorBinary } from './validator-binary.js';

export interface PlatformState {
  suffix: string;
  /** The current binary version is on npm, or is being published in this run. */
  current: boolean;
  /** Versions on npm inside the wrapper's range. */
  inRange: string[];
}

export interface PlatformVerdict {
  ok: boolean;
  message: string;
}

export const platformPackageName = (suffix: string): string => `@attentionform/ooxml-validator-${suffix}`;

/** What is on npm for one platform, judged against the wrapper's binary range. */
export function platformState(suffix: string, binary: ValidatorBinary, versions: readonly string[], publishingNow = false): PlatformState {
  const inRange = versions.filter((v) => inBinaryRange(v, binary.range));
  return { suffix, current: publishingNow || versions.includes(binary.version), inRange };
}

export function judgePlatforms(states: readonly PlatformState[], requireAll: boolean): PlatformVerdict {
  const newest = (s: PlatformState) => [...s.inRange].sort(compareVersions).pop();
  const bare = states.filter((s) => !s.current && s.inRange.length === 0).map((s) => s.suffix);
  if (bare.length > 0) {
    return { ok: false, message: `No binary inside the wrapper's range is on npm for: ${bare.join(', ')}. Users on those platforms would get no validator. Publish those platform packages first.` };
  }
  const behind = states.filter((s) => !s.current);
  if (requireAll && behind.length > 0) {
    return { ok: false, message: `Not on the current binary version: ${behind.map((s) => s.suffix).join(', ')}. This release needs every platform on it (require_all_platforms).` };
  }
  if (behind.length === 0) return { ok: true, message: `All ${states.length} platforms have the current binary.` };
  const list = behind.map((s) => `${s.suffix} (stays on ${newest(s)})`).join(', ');
  return { ok: true, message: `Platforms that keep an older binary inside the range: ${list}.` };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const binary = readValidatorBinary();
  const states = VALIDATOR_TARGETS.map((t) => platformState(t.npmSuffix, binary, publishedVersions(platformPackageName(t.npmSuffix))));
  const verdict = judgePlatforms(states, process.argv.includes('--require-all'));
  console.log(`[validator-platform-status] binary ${binary.version}, range ${binary.range}: ${verdict.message}`);
  if (process.env.GITHUB_STEP_SUMMARY) {
    fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, `### Validator platform packages (binary ${binary.version})\n\n${verdict.message}\n\n`);
  }
  process.exit(verdict.ok ? 0 : 1);
}
