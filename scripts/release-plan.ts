#!/usr/bin/env tsx
/**
 * Works out what a release run would do, before it does anything, and stops it when the answer is
 * wrong. It reads the package manifests and asks npm (read only). Used by .github/workflows/release.yml.
 *
 *   npx tsx scripts/release-plan.ts --selected engine,cli | all [--require-all-platforms] [--check-auth]
 *
 * A package is "published" in a run when it is selected and its version is not on npm yet. The checks:
 *  - something has to be published (a run that would do nothing is a mistake, not a success);
 *  - a package that needs another package of ours must find it on npm, or in the same run;
 *  - the wrapper needs a usable binary on npm for every platform (see validator-platform-status.ts);
 *  - every published package asks for the same Node floor.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { publishedVersions, runNpm, type NpmRunner } from './npm-publish-if-new.js';
import { judgePlatforms, platformPackageName, platformState } from './validator-platform-status.js';
import { VALIDATOR_TARGETS } from './validator-targets.js';
import { compareVersions, parseValidatorBinary, type ValidatorBinary } from './validator-binary.js';

export const UNIT_IDS = ['validator-platforms', 'validator-wrapper', 'engine', 'cli', 'mcp-server'] as const;
export type UnitId = (typeof UNIT_IDS)[number];

/** Repository folder of each single-package unit. Platform packages are built in CI and have none. */
export const UNIT_DIRS: Readonly<Record<Exclude<UnitId, 'validator-platforms'>, string>> = {
  'validator-wrapper': 'packages/ooxml-validator',
  engine: 'packages/engine',
  cli: 'packages/cli',
  'mcp-server': 'packages/mcp-server',
};

export interface Manifest {
  name: string;
  version: string;
  engines?: { node?: string };
  dependencies?: Record<string, string>;
  validatorBinary?: unknown;
}

export interface PlanInput {
  selected: readonly UnitId[];
  requireAllPlatforms: boolean;
  manifests: Readonly<Record<Exclude<UnitId, 'validator-platforms'>, Manifest>>;
  /** Every version of a package on the registry; empty when it does not exist. */
  registryVersions: (name: string) => string[];
}

export interface PlanRow {
  unit: UnitId;
  name: string;
  version: string;
  onNpm: boolean;
  action: 'publish' | 'skip';
}

export interface Plan {
  rows: PlanRow[];
  problems: string[];
  notes: string[];
  /** What will be published, per unit. */
  publish: Record<UnitId, boolean>;
  needsTests: boolean;
}

/** Whether `version` is allowed by `range`, for the two forms our manifests use: `^X.Y.Z` and an exact version. */
export function satisfiesDependency(version: string, range: string): boolean | null {
  const caret = /^\^(\d+)\.(\d+)\.(\d+)$/.exec(range);
  if (caret) {
    const [, major, minor, patch] = caret;
    const floor = `${major}.${minor}.${patch}`;
    if (!/^\d+\.\d+\.\d+$/.test(version) || compareVersions(version, floor) < 0) return false;
    const ceiling = major !== '0' ? `${Number(major) + 1}.0.0` : minor !== '0' ? `0.${Number(minor) + 1}.0` : `0.0.${Number(patch) + 1}`;
    return compareVersions(version, ceiling) < 0;
  }
  if (/^\d+\.\d+\.\d+$/.test(range)) return version === range;
  return null;
}

export function planRelease(input: PlanInput): Plan {
  const { selected, requireAllPlatforms, manifests } = input;
  const cache = new Map<string, string[]>();
  const versionsOf = (name: string) => {
    if (!cache.has(name)) cache.set(name, input.registryVersions(name));
    return cache.get(name) as string[];
  };
  const isSelected = (unit: UnitId) => selected.includes(unit);
  const problems: string[] = [];
  const notes: string[] = [];
  const rows: PlanRow[] = [];
  const publish = Object.fromEntries(UNIT_IDS.map((u) => [u, false])) as Record<UnitId, boolean>;

  if (selected.length === 0) problems.push('Nothing is selected. Choose at least one package.');

  const binary: ValidatorBinary | null = (() => {
    try {
      return parseValidatorBinary(manifests['validator-wrapper'], 'packages/ooxml-validator/package.json');
    } catch (error) {
      problems.push(error instanceof Error ? error.message : String(error));
      return null;
    }
  })();

  const addRow = (unit: UnitId, name: string, version: string) => {
    const onNpm = versionsOf(name).includes(version);
    rows.push({ unit, name, version, onNpm, action: onNpm ? 'skip' : 'publish' });
    if (!onNpm) publish[unit] = true;
  };
  if (isSelected('validator-platforms') && binary) {
    for (const t of VALIDATOR_TARGETS) addRow('validator-platforms', platformPackageName(t.npmSuffix), binary.version);
  }
  for (const unit of ['validator-wrapper', 'engine', 'cli', 'mcp-server'] as const) {
    if (isSelected(unit)) addRow(unit, manifests[unit].name, manifests[unit].version);
  }

  if (selected.length > 0 && !UNIT_IDS.some((u) => publish[u])) {
    problems.push('Every selected version is already on npm, so this run would publish nothing. Raise the version in the package.json files first.');
  }

  // One Node floor for everything we publish.
  const floors = new Set(Object.values(manifests).map((m) => m.engines?.node ?? '(none)'));
  if (floors.size > 1) problems.push(`The packages disagree about the Node version they need: ${[...floors].join(', ')}.`);

  // A package that needs another of ours finds it on npm, or in this run.
  const unitOfName = new Map<string, Exclude<UnitId, 'validator-platforms'>>(
    (Object.keys(UNIT_DIRS) as Array<keyof typeof UNIT_DIRS>).map((u) => [manifests[u].name, u]),
  );
  for (const unit of ['validator-wrapper', 'engine', 'cli', 'mcp-server'] as const) {
    if (!publish[unit]) continue;
    for (const [dep, range] of Object.entries(manifests[unit].dependencies ?? {})) {
      const provider = unitOfName.get(dep);
      if (!provider) continue;
      const inThisRun = publish[provider] ? satisfiesDependency(manifests[provider].version, range) : false;
      const onRegistry = versionsOf(dep).some((v) => satisfiesDependency(v, range) === true);
      if (satisfiesDependency(manifests[provider].version, range) === null) {
        problems.push(`${manifests[unit].name} asks for ${dep} at "${range}", which this check does not understand. Use ^X.Y.Z or an exact version.`);
      } else if (!inThisRun && !onRegistry) {
        problems.push(`${manifests[unit].name} needs ${dep}@${range}, which is not on npm and is not part of this run.`);
      }
    }
  }

  // The wrapper needs a binary for every platform.
  if (publish['validator-wrapper'] && binary) {
    const states = VALIDATOR_TARGETS.map((t) => {
      const name = platformPackageName(t.npmSuffix);
      return platformState(t.npmSuffix, binary, versionsOf(name), publish['validator-platforms'] && !versionsOf(name).includes(binary.version));
    });
    const verdict = judgePlatforms(states, requireAllPlatforms);
    (verdict.ok ? notes : problems).push(verdict.message);
  }

  const needsTests = (['validator-wrapper', 'engine', 'cli', 'mcp-server'] as const).some((u) => publish[u]);
  return { rows, problems, notes, publish, needsTests };
}

export function renderPlan(plan: Plan): string {
  const lines = ['| Package | Version | On npm | This run |', '|---|---|---|---|'];
  for (const r of plan.rows) lines.push(`| ${r.name} | ${r.version} | ${r.onNpm ? 'yes' : 'no'} | ${r.action === 'publish' ? '**publish**' : 'skip (already on npm)'} |`);
  if (plan.rows.length === 0) lines.push('| (nothing selected) | | | |');
  const out = [lines.join('\n')];
  if (plan.notes.length) out.push('', ...plan.notes.map((n) => `- ${n}`));
  if (plan.problems.length) out.push('', '**Stopped:**', ...plan.problems.map((p) => `- ${p}`));
  return out.join('\n');
}

export function parseSelection(text: string): UnitId[] {
  const parts = text.split(',').map((p) => p.trim()).filter(Boolean);
  if (parts.includes('all')) return [...UNIT_IDS];
  const unknown = parts.filter((p) => !(UNIT_IDS as readonly string[]).includes(p));
  if (unknown.length > 0) throw new Error(`Unknown package selection: ${unknown.join(', ')}. Choose from ${UNIT_IDS.join(', ')} or all.`);
  return UNIT_IDS.filter((u) => parts.includes(u));
}

/** Confirms the npm token logs in. Only a login is checked: npm has no read-only way to ask about publish rights. */
export function checkLogin(run: NpmRunner = runNpm): string {
  const r = run(['whoami']);
  if (r.status !== 0) throw new Error(`npm login failed (${r.stderr.trim() || `exit ${r.status}`}). Check the NPM_TOKEN secret.`);
  return r.stdout.trim();
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const args = process.argv.slice(2);
  const value = (flag: string) => {
    const at = args.indexOf(flag);
    return at >= 0 ? (args[at + 1] ?? '') : '';
  };
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const manifests = Object.fromEntries(
    (Object.entries(UNIT_DIRS) as Array<[keyof typeof UNIT_DIRS, string]>).map(([unit, dir]) => [
      unit,
      JSON.parse(fs.readFileSync(path.join(root, dir, 'package.json'), 'utf8')) as Manifest,
    ]),
  ) as PlanInput['manifests'];

  const plan = planRelease({
    selected: parseSelection(value('--selected')),
    requireAllPlatforms: args.includes('--require-all-platforms'),
    manifests,
    registryVersions: (name) => publishedVersions(name),
  });

  if (args.includes('--check-auth')) {
    if (process.env.NODE_AUTH_TOKEN) {
      try {
        plan.notes.push(`npm login works (as ${checkLogin()}).`);
      } catch (error) {
        plan.problems.push(error instanceof Error ? error.message : String(error));
      }
    } else {
      plan.notes.push('No npm token in this run, so the login was not checked.');
    }
  }

  const text = renderPlan(plan);
  console.log(text);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, `### Release plan\n\n${text}\n\n`);
  if (process.env.GITHUB_OUTPUT) {
    const out = [
      `platforms=${plan.publish['validator-platforms']}`,
      `wrapper=${plan.publish['validator-wrapper']}`,
      `engine=${plan.publish.engine}`,
      `cli=${plan.publish.cli}`,
      `mcp_server=${plan.publish['mcp-server']}`,
      `needs_tests=${plan.needsTests}`,
      `selected=${plan.rows.length > 0 ? UNIT_IDS.filter((u) => plan.publish[u]).join(',') : ''}`,
    ];
    fs.appendFileSync(process.env.GITHUB_OUTPUT, out.join('\n') + '\n');
  }
  process.exit(plan.problems.length > 0 ? 1 : 0);
}
