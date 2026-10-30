#!/usr/bin/env tsx
/**
 * Publishes the packages a release run selected, in dependency order, and tags what it published.
 * Used by .github/workflows/release.yml after the plan (release-plan.ts) and the tests.
 *
 *   npx tsx scripts/release-publish.ts [--platforms] [--wrapper] [--engine] [--cli] [--mcp-server]
 *       [--dry-run] [--provenance] [--require-all-platforms] [--tag] [--platform-dir platform-out]
 *
 * Order: the six platform packages, the wrapper, the engine, the CLI and the MCP server.
 *  - A platform whose package was not built, or whose publish failed, does not stop the others.
 *  - Before the wrapper goes out, every platform must have a binary inside the wrapper's range on npm
 *    (the one just published, or an older one). Otherwise the wrapper is not published.
 *  - Any other failure stops the run: what comes after depends on it.
 *  - A version already on npm is skipped, so a rerun finishes a partial release.
 *  - With --tag, each package published in this run gets a git tag on this commit. Platform packages
 *    are tagged only once all six are on npm at the binary version.
 * The process exits 1 when anything failed, after doing everything it was still allowed to do.
 */
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { publishIfNew, publishedVersions, type PublishOutcome } from './npm-publish-if-new.js';
import { injectPlatformDependencies } from './inject-validator-optional-deps.js';
import { judgePlatforms, platformPackageName, platformState } from './validator-platform-status.js';
import { VALIDATOR_TARGETS } from './validator-targets.js';
import { readValidatorBinary } from './validator-binary.js';
import { UNIT_DIRS, type UnitId } from './release-plan.js';

export interface PublishDeps {
  publish(dir: string, options: { dryRun: boolean; provenance: boolean }): PublishOutcome;
  registryVersions(name: string): string[];
  exists(target: string): boolean;
  /** True when the file is there and has an execute bit (the platform job's artifact upload drops it). */
  isExecutable(file: string): boolean;
  readPackage(dir: string): { name: string; version: string };
  /** Adds the platform packages to the wrapper's manifest; returns a function that puts the manifest back. */
  injectPlatforms(): () => void;
  git(args: string[]): { status: number; stdout: string; stderr: string };
  log(message: string): void;
}

export interface PublishOptions {
  units: Readonly<Record<UnitId, boolean>>;
  dryRun: boolean;
  provenance: boolean;
  requireAllPlatforms: boolean;
  tag: boolean;
  /** Folder holding one built package per platform, named like the platform (linux-x64). */
  platformDir: string;
  /** Repository root, where packages/ lives. */
  root: string;
}

export interface PublishEntry {
  unit: UnitId;
  name: string;
  version: string;
  outcome: PublishOutcome | 'missing' | 'failed' | 'blocked';
  detail?: string;
}

export interface PublishReport {
  entries: PublishEntry[];
  failures: string[];
  tags: string[];
}

export const TAG_PREFIX: Readonly<Record<UnitId, string>> = {
  'validator-platforms': 'validator-binary',
  'validator-wrapper': 'validator',
  engine: 'engine',
  cli: 'cli',
  'mcp-server': 'mcp-server',
};

export const defaultDeps: PublishDeps = {
  publish: (dir, options) => publishIfNew({ dir, ...options }),
  registryVersions: (name) => publishedVersions(name),
  exists: (target) => fs.existsSync(target),
  isExecutable: (file) => fs.existsSync(file) && (fs.statSync(file).mode & 0o111) !== 0,
  readPackage: (dir) => JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8')) as { name: string; version: string },
  injectPlatforms: () => {
    const manifest = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../packages/ooxml-validator/package.json');
    const original = fs.readFileSync(manifest, 'utf8');
    injectPlatformDependencies(manifest);
    return () => fs.writeFileSync(manifest, original, 'utf8');
  },
  git: (args) => {
    const r = spawnSync('git', args, { encoding: 'utf8' });
    return { status: r.status ?? 1, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
  },
  log: (message) => console.log(message),
};

const SINGLE_UNITS = ['validator-wrapper', 'engine', 'cli', 'mcp-server'] as const;

export function runRelease(options: PublishOptions, deps: PublishDeps = defaultDeps): PublishReport {
  const { units, dryRun, provenance, root } = options;
  const entries: PublishEntry[] = [];
  const failures: string[] = [];
  const binary = readValidatorBinary(path.join(root, UNIT_DIRS['validator-wrapper'], 'package.json'));
  const flags = { dryRun, provenance };

  // 1. Platform packages. One failing does not hold back the others.
  const platformOutcome = new Map<string, PublishEntry['outcome']>();
  if (units['validator-platforms']) {
    for (const target of VALIDATOR_TARGETS) {
      const name = platformPackageName(target.npmSuffix);
      const dir = path.join(options.platformDir, target.npmSuffix);
      let entry: PublishEntry;
      if (!deps.exists(dir)) {
        entry = { unit: 'validator-platforms', name, version: binary.version, outcome: 'missing', detail: `no built package at ${dir} (its build job failed or was cancelled)` };
        failures.push(`${name}: not built`);
      } else {
        try {
          // A binary without its execute bit installs fine and fails for every user (EACCES), so refuse it.
          const file = path.join(dir, target.binaryName);
          if (target.os !== 'win32' && !deps.isExecutable(file)) {
            throw new Error(`${file} is missing or not executable; the package would install and then fail with EACCES`);
          }
          entry = { unit: 'validator-platforms', name, version: binary.version, outcome: deps.publish(dir, flags) };
        } catch (error) {
          const detail = error instanceof Error ? error.message : String(error);
          entry = { unit: 'validator-platforms', name, version: binary.version, outcome: 'failed', detail };
          failures.push(`${name}: ${detail}`);
        }
      }
      platformOutcome.set(target.npmSuffix, entry.outcome);
      entries.push(entry);
      deps.log(`[release] ${name}@${binary.version}: ${entry.outcome}${entry.detail ? ` (${entry.detail})` : ''}`);
    }
  }

  // 2. The wrapper, 3. the packages that use it. The first failure here ends the run.
  let stopped = false;
  for (const unit of SINGLE_UNITS) {
    if (!units[unit] || stopped) continue;
    const dir = path.join(root, UNIT_DIRS[unit]);
    const { name, version } = deps.readPackage(dir);
    try {
      if (unit === 'validator-wrapper') {
        const states = VALIDATOR_TARGETS.map((t) => {
          const outcome = platformOutcome.get(t.npmSuffix);
          const justNow = outcome === 'published' || outcome === 'dry-run';
          return platformState(t.npmSuffix, binary, deps.registryVersions(platformPackageName(t.npmSuffix)), justNow);
        });
        const verdict = judgePlatforms(states, options.requireAllPlatforms);
        deps.log(`[release] platforms: ${verdict.message}`);
        if (!verdict.ok) {
          entries.push({ unit, name, version, outcome: 'blocked', detail: verdict.message });
          failures.push(`${name}: not published. ${verdict.message}`);
          stopped = true;
          continue;
        }
      }
      // The wrapper goes out with the platform packages listed; the checkout gets its manifest back.
      const restore = unit === 'validator-wrapper' ? deps.injectPlatforms() : () => undefined;
      let outcome: PublishOutcome;
      try {
        outcome = deps.publish(dir, flags);
      } finally {
        restore();
      }
      entries.push({ unit, name, version, outcome });
      deps.log(`[release] ${name}@${version}: ${outcome}`);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      entries.push({ unit, name, version, outcome: 'failed', detail });
      failures.push(`${name}: ${detail}`);
      stopped = true;
    }
  }

  // 4. Tags for what went out in this run.
  const tags: string[] = [];
  if (options.tag && !dryRun) {
    const toTag: string[] = [];
    for (const unit of SINGLE_UNITS) {
      const e = entries.find((x) => x.unit === unit && x.outcome === 'published');
      if (e) toTag.push(`${TAG_PREFIX[unit]}-v${e.version}`);
    }
    const platforms = entries.filter((e) => e.unit === 'validator-platforms');
    if (platforms.some((e) => e.outcome === 'published')) {
      const allThere = VALIDATOR_TARGETS.every((t) => {
        const o = platformOutcome.get(t.npmSuffix);
        return o === 'published' || o === 'skipped';
      });
      if (allThere) toTag.push(`${TAG_PREFIX['validator-platforms']}-v${binary.version}`);
      else deps.log('[release] not tagging the binary version: not every platform is on npm at it');
    }
    for (const tag of toTag) {
      const existing = deps.git(['rev-parse', '-q', '--verify', `refs/tags/${tag}`]);
      if (existing.status === 0) {
        deps.log(`[release] tag ${tag} already exists, left as it is`);
        continue;
      }
      const made = deps.git(['tag', tag]);
      const pushed = made.status === 0 ? deps.git(['push', 'origin', `refs/tags/${tag}`]) : made;
      if (pushed.status === 0) tags.push(tag);
      else failures.push(`tag ${tag} could not be created or pushed: ${pushed.stderr.trim()}`);
    }
  }

  return { entries, failures, tags };
}

export function renderReport(report: PublishReport, dryRun: boolean): string {
  const lines = [`| Package | Version | Result |`, `|---|---|---|`];
  for (const e of report.entries) lines.push(`| ${e.name} | ${e.version} | ${e.outcome}${e.detail ? ` — ${e.detail}` : ''} |`);
  if (dryRun) lines.push('', 'Dry run: nothing was published and nothing was tagged.');
  if (report.tags.length) lines.push('', `Tags: ${report.tags.join(', ')}`);
  if (report.failures.length) lines.push('', '**Failed:**', ...report.failures.map((f) => `- ${f}`));
  return lines.join('\n');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const args = process.argv.slice(2);
  const value = (flag: string, fallback: string) => {
    const at = args.indexOf(flag);
    return at >= 0 ? (args[at + 1] ?? fallback) : fallback;
  };
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const dryRun = args.includes('--dry-run');
  const report = runRelease({
    units: {
      'validator-platforms': args.includes('--platforms'),
      'validator-wrapper': args.includes('--wrapper'),
      engine: args.includes('--engine'),
      cli: args.includes('--cli'),
      'mcp-server': args.includes('--mcp-server'),
    },
    dryRun,
    provenance: args.includes('--provenance'),
    requireAllPlatforms: args.includes('--require-all-platforms'),
    tag: args.includes('--tag'),
    platformDir: path.resolve(value('--platform-dir', 'platform-out')),
    root,
  });
  const text = renderReport(report, dryRun);
  console.log(text);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, `### ${dryRun ? 'Dry run' : 'Release'} result\n\n${text}\n\n`);
  process.exit(report.failures.length > 0 ? 1 : 0);
}
