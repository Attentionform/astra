#!/usr/bin/env tsx
/**
 * Publishes the package in a directory unless that exact version is already on npm, then checks
 * that the registry serves it. A rerun after a partial failure can therefore finish a release
 * without failing on the packages that were already published.
 *
 *   npx tsx scripts/npm-publish-if-new.ts <dir> [--dry-run] [--provenance]
 */
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export interface NpmResult {
  status: number;
  stdout: string;
  stderr: string;
}
export type NpmRunner = (args: string[]) => NpmResult;

export const runNpm: NpmRunner = (args) => {
  // npm is npm.cmd on Windows, which Node only starts through a shell.
  const r = spawnSync('npm', args, { encoding: 'utf8', shell: process.platform === 'win32' });
  return { status: r.status ?? 1, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
};

/** True when `name@version` is on the registry, false when it is not; throws when the registry cannot be asked. */
export function isPublished(name: string, version: string, run: NpmRunner = runNpm): boolean {
  const r = run(['view', `${name}@${version}`, 'version']);
  if (r.status === 0) return r.stdout.trim() === version;
  if (r.stderr.includes('E404')) return false;
  throw new Error(`Could not ask the registry about ${name}@${version}: ${r.stderr.trim() || `exit ${r.status}`}`);
}

/** Every version of `name` on the registry; an empty list when the package does not exist. Throws when the registry cannot be asked. */
export function publishedVersions(name: string, run: NpmRunner = runNpm): string[] {
  const r = run(['view', name, 'versions', '--json']);
  if (r.status !== 0) {
    if (r.stderr.includes('E404')) return [];
    throw new Error(`Could not ask the registry about ${name}: ${r.stderr.trim() || `exit ${r.status}`}`);
  }
  const text = r.stdout.trim();
  if (text === '') return [];
  const parsed: unknown = JSON.parse(text);
  return (Array.isArray(parsed) ? parsed : [parsed]).filter((v): v is string => typeof v === 'string');
}

export type PublishOutcome = 'published' | 'skipped' | 'dry-run';

/**
 * After a publish succeeds, npm still has to process the new version before it serves it. For a new
 * package that took about 1 minute (a tiny package) to 3 minutes (a 35 MB binary), so ask every 10
 * seconds for up to 10 minutes. The loop ends the moment the version is visible.
 */
export const VISIBILITY_ATTEMPTS = 60;
export const VISIBILITY_WAIT_MS = 10_000;

export interface PublishOptions {
  dir: string;
  dryRun?: boolean;
  provenance?: boolean;
  run?: NpmRunner;
  /** How many times to ask the registry whether a new version is visible, one `wait` apart. */
  attempts?: number;
  wait?: () => void;
}

export function publishIfNew(options: PublishOptions): PublishOutcome {
  const { dir, dryRun = false, provenance = false, run = runNpm, attempts = VISIBILITY_ATTEMPTS, wait = () => sleepSync(VISIBILITY_WAIT_MS) } = options;
  const pkg = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8')) as { name: string; version: string };

  if (!dryRun && isPublished(pkg.name, pkg.version, run)) {
    console.log(`[npm-publish-if-new] ${pkg.name}@${pkg.version} is already on npm, skipping`);
    return 'skipped';
  }

  const args = ['publish', dir, '--access', 'public'];
  if (provenance && !dryRun) args.push('--provenance');
  if (dryRun) args.push('--dry-run');
  const r = run(args);
  if (r.status !== 0) throw new Error(`npm publish failed for ${pkg.name}@${pkg.version}: ${r.stderr.trim() || r.stdout.trim()}`);
  if (dryRun) {
    console.log(`[npm-publish-if-new] dry run: ${pkg.name}@${pkg.version} would be published`);
    return 'dry-run';
  }

  for (let i = 1; i <= attempts; i++) {
    if (isPublished(pkg.name, pkg.version, run)) {
      console.log(`[npm-publish-if-new] published ${pkg.name}@${pkg.version}`);
      return 'published';
    }
    if (i < attempts) wait();
  }
  throw new Error(`${pkg.name}@${pkg.version} was published but the registry does not serve it yet`);
}

function sleepSync(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const args = process.argv.slice(2);
  const dir = args.find((a) => !a.startsWith('--'));
  if (!dir) {
    console.error('usage: npx tsx scripts/npm-publish-if-new.ts <dir> [--dry-run] [--provenance]');
    process.exit(2);
  }
  publishIfNew({ dir: path.resolve(dir), dryRun: args.includes('--dry-run'), provenance: args.includes('--provenance') });
}
