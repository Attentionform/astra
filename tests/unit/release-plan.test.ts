/**
 * What a release run decides before it publishes anything: which packages go out, and every reason
 * to stop. The registry is faked; the manifests are the repository's own.
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import {
  UNIT_DIRS,
  UNIT_IDS,
  checkLogin,
  parseSelection,
  planRelease,
  renderPlan,
  satisfiesDependency,
  type Manifest,
  type PlanInput,
  type UnitId,
} from '../../scripts/release-plan.js';
import { VALIDATOR_TARGETS } from '../../scripts/validator-targets.js';

const ROOT = path.resolve(__dirname, '../..');
const manifests = Object.fromEntries(
  Object.entries(UNIT_DIRS).map(([unit, dir]) => [unit, JSON.parse(fs.readFileSync(path.join(ROOT, dir, 'package.json'), 'utf8')) as Manifest]),
) as PlanInput['manifests'];

type Registry = Record<string, string[]>;
const platformName = (suffix: string) => `@attentionform/ooxml-validator-${suffix}`;
const allPlatforms = (versions: string[]): Registry => Object.fromEntries(VALIDATOR_TARGETS.map((t) => [platformName(t.npmSuffix), versions]));
const plan = (selected: UnitId[], registry: Registry = {}, extra: Partial<PlanInput> = {}) =>
  planRelease({ selected, requireAllPlatforms: false, manifests, registryVersions: (name) => registry[name] ?? [], ...extra });

describe('the first release: nothing is on npm', () => {
  it('publishes everything when everything is selected, platforms first', () => {
    const p = plan([...UNIT_IDS]);
    expect(p.problems).toEqual([]);
    expect(UNIT_IDS.every((u) => p.publish[u])).toBe(true);
    expect(p.rows.filter((r) => r.unit === 'validator-platforms')).toHaveLength(6);
    expect(p.needsTests).toBe(true);
  });

  it('stops a wrapper release that has no binary to point to', () => {
    const p = plan(['validator-wrapper', 'engine', 'cli', 'mcp-server']);
    expect(p.problems.join('\n')).toMatch(/No binary inside the wrapper's range is on npm for: linux-x64/);
  });

  it('stops the CLI when the engine it needs is nowhere', () => {
    const p = plan(['cli'], { '@attentionform/ooxml-validator': ['0.1.0'] });
    expect(p.problems.join('\n')).toContain('@attentionform/astra needs @attentionform/astra-engine@^0.1.0');
  });

  it('lets the CLI go together with the engine it needs', () => {
    const p = plan(['engine', 'cli'], { '@attentionform/ooxml-validator': ['0.1.0'] });
    expect(p.problems).toEqual([]);
  });
});

describe('releasing one thing', () => {
  const published: Registry = {
    '@attentionform/ooxml-validator': ['0.1.0'],
    '@attentionform/astra-engine': ['0.1.0'],
    ...allPlatforms(['0.1.0']),
  };

  it('can publish the engine alone when what it needs is on npm', () => {
    const p = plan(['engine'], { ...published, '@attentionform/astra-engine': [] });
    expect(p.problems).toEqual([]);
    expect(p.publish.engine).toBe(true);
    expect(p.publish.cli).toBe(false);
    expect(p.publish['validator-platforms']).toBe(false);
  });

  it('can publish the platforms alone, with no tests needed', () => {
    const p = plan(['validator-platforms'], { ...published, ...allPlatforms([]) });
    expect(p.problems).toEqual([]);
    expect(p.needsTests).toBe(false);
  });

  it('skips a version that is already on npm and says so', () => {
    const p = plan(['engine', 'cli'], { ...published, '@attentionform/astra': [] });
    expect(p.publish.engine).toBe(false);
    expect(p.publish.cli).toBe(true);
    expect(renderPlan(p)).toContain('skip (already on npm)');
  });

  it('refuses a run that would publish nothing', () => {
    const p = plan(['engine'], published);
    expect(p.problems.join('\n')).toMatch(/would publish nothing/);
  });

  it('refuses an empty selection', () => {
    expect(plan([]).problems.join('\n')).toContain('Nothing is selected');
  });
});

describe('the wrapper and its platform binaries', () => {
  const base: Registry = { '@attentionform/ooxml-validator': [] };

  it('stops the wrapper when a platform only has a binary outside the range', () => {
    const p = plan(['validator-wrapper'], { ...base, ...allPlatforms(['0.1.0']), [platformName('linux-arm64')]: ['0.0.9'] });
    expect(p.problems.join('\n')).toContain('linux-arm64');
  });

  it('with the range moved up, an older platform binary keeps the wrapper usable', () => {
    const moved = { ...manifests, 'validator-wrapper': { ...manifests['validator-wrapper'], validatorBinary: { version: '0.1.2', range: '>=0.1.0 <0.2.0' } } };
    const p = planRelease({
      selected: ['validator-wrapper'],
      requireAllPlatforms: false,
      manifests: moved,
      registryVersions: (name) => ({ ...base, ...allPlatforms(['0.1.0', '0.1.2']), [platformName('win32-arm64')]: ['0.1.0'] })[name] ?? [],
    });
    expect(p.problems).toEqual([]);
    expect(p.notes.join('\n')).toContain('win32-arm64 (stays on 0.1.0)');
  });

  it('with require-all, stops a wrapper whose platform is on an older binary', () => {
    const moved = { ...manifests, 'validator-wrapper': { ...manifests['validator-wrapper'], validatorBinary: { version: '0.1.2', range: '>=0.1.0 <0.2.0' } } };
    const p = planRelease({
      selected: ['validator-wrapper'],
      requireAllPlatforms: true,
      manifests: moved,
      registryVersions: (name) => ({ ...base, ...allPlatforms(['0.1.2']), [platformName('darwin-x64')]: ['0.1.0'] })[name] ?? [],
    });
    expect(p.problems.join('\n')).toContain('darwin-x64');
  });

  it('counts platforms published in the same run', () => {
    const p = plan(['validator-platforms', 'validator-wrapper'], base, { requireAllPlatforms: true });
    expect(p.problems).toEqual([]);
  });

  it('refuses a manifest whose binary version is outside its own range', () => {
    const broken = { ...manifests, 'validator-wrapper': { ...manifests['validator-wrapper'], validatorBinary: { version: '0.3.0', range: '>=0.1.0 <0.2.0' } } };
    const p = planRelease({ selected: ['validator-platforms'], requireAllPlatforms: false, manifests: broken, registryVersions: () => [] });
    expect(p.problems.join('\n')).toContain('outside its own range');
  });
});

describe('consistency', () => {
  it('stops when the packages disagree about the Node version', () => {
    const odd = { ...manifests, cli: { ...manifests.cli, engines: { node: '>=20.0.0' } } };
    const p = planRelease({ selected: ['cli'], requireAllPlatforms: false, manifests: odd, registryVersions: () => [] });
    expect(p.problems.join('\n')).toMatch(/disagree about the Node version/);
  });

  it('stops on a dependency range it does not understand instead of guessing', () => {
    const odd = { ...manifests, cli: { ...manifests.cli, dependencies: { ...manifests.cli.dependencies, '@attentionform/astra-engine': '>=0.1.0' } } };
    const p = planRelease({ selected: ['cli'], requireAllPlatforms: false, manifests: odd, registryVersions: () => [] });
    expect(p.problems.join('\n')).toContain('does not understand');
  });
});

describe('satisfiesDependency', () => {
  it.each([
    ['0.1.0', '^0.1.0', true],
    ['0.1.9', '^0.1.0', true],
    ['0.2.0', '^0.1.0', false],
    ['0.0.9', '^0.1.0', false],
    ['1.4.0', '^1.2.0', true],
    ['2.0.0', '^1.2.0', false],
    ['0.0.3', '^0.0.3', true],
    ['0.0.4', '^0.0.3', false],
    ['1.0.0', '1.0.0', true],
    ['1.0.1', '1.0.0', false],
    ['0.1.0-beta.1', '^0.1.0', false],
  ])('%s against %s is %s', (version, range, expected) => expect(satisfiesDependency(version, range)).toBe(expected));
  it('says it does not know a range it cannot read', () => expect(satisfiesDependency('1.0.0', '>=1.0.0')).toBeNull());
});

describe('selection and login', () => {
  it('reads "all" and keeps the publish order', () => {
    expect(parseSelection('all')).toEqual([...UNIT_IDS]);
    expect(parseSelection('cli, engine')).toEqual(['engine', 'cli']);
  });
  it('refuses a name it does not know', () => expect(() => parseSelection('engine,nope')).toThrow(/Unknown package selection: nope/));
  it('reports who npm says is logged in', () => expect(checkLogin(() => ({ status: 0, stdout: 'releaser\n', stderr: '' }))).toBe('releaser'));
  it('fails with a clear message when the login does not work', () => {
    expect(() => checkLogin(() => ({ status: 1, stdout: '', stderr: 'npm error code E401' }))).toThrow(/NPM_TOKEN/);
  });
});
