/**
 * The order things are published in, what keeps going when a platform fails, what stops a run, and
 * what gets tagged. Registry, git and npm are faked; the manifests are the repository's own.
 */
import path from 'node:path';
import { describe, it, expect } from 'vitest';
import { runRelease, renderReport, type PublishDeps, type PublishOptions } from '../../scripts/release-publish.js';
import type { PublishOutcome } from '../../scripts/npm-publish-if-new.js';
import { VALIDATOR_TARGETS } from '../../scripts/validator-targets.js';

const ROOT = path.resolve(__dirname, '../..');
const PLATFORM_DIR = path.join(ROOT, 'platform-out');
const NAMES: Record<string, { name: string; version: string }> = {
  'packages/ooxml-validator': { name: '@attentionform/ooxml-validator', version: '0.1.0' },
  'packages/engine': { name: '@attentionform/astra-engine', version: '0.1.0' },
  'packages/cli': { name: '@attentionform/astra', version: '0.1.0' },
  'packages/mcp-server': { name: '@attentionform/astra-mcp-server', version: '0.1.0' },
};
const platformName = (suffix: string) => `@attentionform/ooxml-validator-${suffix}`;

interface Fake {
  deps: PublishDeps;
  published: string[];
  injected: () => number;
  restored: () => number;
  gitCalls: string[][];
  logs: string[];
}

function fake(
  config: {
    missingDirs?: string[];
    outcomes?: Record<string, PublishOutcome | Error>;
    registry?: Record<string, string[]>;
    existingTags?: string[];
    gitFails?: string;
    notExecutable?: string[];
  } = {},
): Fake {
  const published: string[] = [];
  const gitCalls: string[][] = [];
  const logs: string[] = [];
  let injected = 0;
  let restored = 0;
  const deps: PublishDeps = {
    publish: (dir, options) => {
      const key = path.relative(ROOT, dir).replace(/^platform-out\//, 'platform:');
      published.push(options.dryRun ? `${key} (dry)` : key);
      const outcome = config.outcomes?.[key];
      if (outcome instanceof Error) throw outcome;
      return outcome ?? (options.dryRun ? 'dry-run' : 'published');
    },
    registryVersions: (name) => config.registry?.[name] ?? [],
    exists: (target) => !(config.missingDirs ?? []).some((d) => target === path.join(PLATFORM_DIR, d)),
    isExecutable: (file) => !(config.notExecutable ?? []).some((d) => file === path.join(PLATFORM_DIR, d, 'ooxml-validator')),
    readPackage: (dir) => NAMES[path.relative(ROOT, dir)],
    injectPlatforms: () => {
      injected++;
      return () => void restored++;
    },
    git: (args) => {
      gitCalls.push(args);
      if (args[0] === 'rev-parse') return { status: (config.existingTags ?? []).some((t) => args[3] === `refs/tags/${t}`) ? 0 : 1, stdout: '', stderr: '' };
      if (config.gitFails && args[0] === config.gitFails) return { status: 1, stdout: '', stderr: `${args[0]} refused` };
      return { status: 0, stdout: '', stderr: '' };
    },
    log: (m) => logs.push(m),
  };
  return { deps, published, injected: () => injected, restored: () => restored, gitCalls, logs };
}

const everything = { 'validator-platforms': true, 'validator-wrapper': true, engine: true, cli: true, 'mcp-server': true } as const;
const nothing = { 'validator-platforms': false, 'validator-wrapper': false, engine: false, cli: false, 'mcp-server': false } as const;
const options = (over: Partial<PublishOptions> = {}): PublishOptions => ({
  units: everything,
  dryRun: false,
  provenance: false,
  requireAllPlatforms: false,
  tag: true,
  platformDir: PLATFORM_DIR,
  root: ROOT,
  ...over,
});
const platformKeys = VALIDATOR_TARGETS.map((t) => `platform:${t.npmSuffix}`);
const olderBinaries = Object.fromEntries(VALIDATOR_TARGETS.map((t) => [platformName(t.npmSuffix), ['0.1.0']]));

describe('a full release', () => {
  it('publishes the platforms, then the wrapper, the engine, the CLI and the MCP server', () => {
    const f = fake();
    const report = runRelease(options(), f.deps);
    expect(f.published).toEqual([...platformKeys, 'packages/ooxml-validator', 'packages/engine', 'packages/cli', 'packages/mcp-server']);
    expect(report.failures).toEqual([]);
    expect(f.injected()).toBe(1);
  });

  it('tags what it published, once the six platforms are all on npm', () => {
    const f = fake();
    const report = runRelease(options(), f.deps);
    expect(report.tags.sort()).toEqual(['cli-v0.1.0', 'engine-v0.1.0', 'mcp-server-v0.1.0', 'validator-binary-v0.1.0', 'validator-v0.1.0']);
    expect(f.gitCalls.filter((c) => c[0] === 'push')).toHaveLength(5);
  });

  it('a dry run publishes and tags nothing, but runs every step so it shows what would happen', () => {
    const f = fake();
    const report = runRelease(options({ dryRun: true }), f.deps);
    expect(f.published.every((p) => p.endsWith('(dry)'))).toBe(true);
    expect(f.published).toHaveLength(10);
    expect(report.tags).toEqual([]);
    expect(f.gitCalls).toEqual([]);
    expect(renderReport(report, true)).toContain('Dry run: nothing was published');
  });
});

describe('a platform that failed', () => {
  it('does not stop the other platforms or, with an older binary in range on npm, the wrapper', () => {
    const f = fake({ missingDirs: ['linux-arm64'], registry: olderBinaries });
    const report = runRelease(options(), f.deps);
    expect(f.published).toContain('packages/ooxml-validator');
    expect(f.published).toContain('packages/engine');
    expect(report.failures).toHaveLength(1);
    expect(report.failures[0]).toContain('linux-arm64');
    expect(report.entries.find((e) => e.name.endsWith('linux-arm64'))?.outcome).toBe('missing');
  });

  it('does not tag the binary version when a platform is not on npm at it', () => {
    const f = fake({ missingDirs: ['linux-arm64'], registry: olderBinaries });
    const report = runRelease(options(), f.deps);
    expect(report.tags).not.toContain('validator-binary-v0.1.0');
    expect(report.tags).toContain('validator-v0.1.0');
  });

  it('is not published when its binary has lost the execute bit, and the others still go out', () => {
    const f = fake({ notExecutable: ['darwin-arm64'], registry: olderBinaries });
    const report = runRelease(options(), f.deps);
    expect(f.published).not.toContain('platform:darwin-arm64');
    expect(f.published.filter((p) => p.startsWith('platform:'))).toHaveLength(5);
    const entry = report.entries.find((e) => e.name.endsWith('darwin-arm64'));
    expect(entry?.outcome).toBe('failed');
    expect(entry?.detail).toContain('not executable');
    expect(entry?.detail).toContain('EACCES');
  });

  it('keeps going when a platform publish throws', () => {
    const f = fake({ outcomes: { 'platform:win32-x64': new Error('E403 forbidden') }, registry: olderBinaries });
    const report = runRelease(options(), f.deps);
    expect(f.published.filter((p) => p.startsWith('platform:'))).toHaveLength(6);
    expect(report.failures.join('\n')).toContain('E403 forbidden');
    expect(f.published).toContain('packages/ooxml-validator');
  });

  it('holds the wrapper back when a platform has no binary in range at all, and publishes nothing after it', () => {
    const f = fake({ missingDirs: ['linux-arm64'], registry: { ...olderBinaries, [platformName('linux-arm64')]: [] } });
    const report = runRelease(options(), f.deps);
    expect(f.published).not.toContain('packages/ooxml-validator');
    expect(f.published).not.toContain('packages/engine');
    expect(f.injected()).toBe(0);
    expect(report.entries.find((e) => e.unit === 'validator-wrapper')?.outcome).toBe('blocked');
    expect(report.failures.join('\n')).toContain('No binary inside the wrapper');
  });

  it('with require-all, holds the wrapper back when a platform is only on an older binary', () => {
    const f = fake({ missingDirs: ['linux-arm64'], registry: olderBinaries });
    // the registry has 0.1.0 and the binary being released is 0.1.0, so make the platform miss it
    const report = runRelease(options({ requireAllPlatforms: true }), { ...f.deps, registryVersions: () => ['0.0.9'] });
    expect(report.entries.find((e) => e.unit === 'validator-wrapper')?.outcome).toBe('blocked');
  });
});

describe('the wrapper manifest', () => {
  it('is put back after the wrapper is published', () => {
    const f = fake();
    runRelease(options({ units: { ...nothing, 'validator-wrapper': true }, requireAllPlatforms: false }), { ...f.deps, registryVersions: () => ['0.1.0'] });
    expect(f.injected()).toBe(1);
    expect(f.restored()).toBe(1);
  });

  it('is put back even when the publish fails', () => {
    const f = fake({ outcomes: { 'packages/ooxml-validator': new Error('E403') } });
    runRelease(options({ units: { ...nothing, 'validator-wrapper': true } }), { ...f.deps, registryVersions: () => ['0.1.0'] });
    expect(f.restored()).toBe(1);
  });
});

describe('what stops a run', () => {
  it('a failing wrapper keeps the engine, CLI and MCP server from going out', () => {
    const f = fake({ outcomes: { 'packages/ooxml-validator': new Error('E403 forbidden') } });
    const report = runRelease(options(), f.deps);
    expect(f.published).not.toContain('packages/engine');
    expect(report.failures.join('\n')).toContain('E403');
  });

  it('a failing engine keeps the CLI and MCP server back, but the wrapper stays published and tagged', () => {
    const f = fake({ outcomes: { 'packages/engine': new Error('E500') } });
    const report = runRelease(options(), f.deps);
    expect(f.published).not.toContain('packages/cli');
    expect(report.tags).toContain('validator-v0.1.0');
    expect(report.tags).not.toContain('engine-v0.1.0');
  });
});

describe('releasing one thing', () => {
  it('the engine alone touches nothing else', () => {
    const f = fake();
    const report = runRelease(options({ units: { ...nothing, engine: true } }), f.deps);
    expect(f.published).toEqual(['packages/engine']);
    expect(f.injected()).toBe(0);
    expect(report.tags).toEqual(['engine-v0.1.0']);
  });

  it('the platforms alone publish six packages and tag the binary version', () => {
    const f = fake();
    const report = runRelease(options({ units: { ...nothing, 'validator-platforms': true } }), f.deps);
    expect(f.published).toEqual(platformKeys);
    expect(report.tags).toEqual(['validator-binary-v0.1.0']);
  });

  it('a rerun that finds everything already on npm tags nothing', () => {
    const f = fake({ outcomes: Object.fromEntries([...platformKeys, 'packages/ooxml-validator', 'packages/engine', 'packages/cli', 'packages/mcp-server'].map((k) => [k, 'skipped' as const])), registry: olderBinaries });
    const report = runRelease(options(), f.deps);
    expect(report.tags).toEqual([]);
    expect(report.failures).toEqual([]);
  });
});

describe('tags', () => {
  it('leaves a tag that already exists alone', () => {
    const f = fake({ existingTags: ['engine-v0.1.0'] });
    const report = runRelease(options({ units: { ...nothing, engine: true } }), f.deps);
    expect(report.tags).toEqual([]);
    expect(f.gitCalls.some((c) => c[0] === 'tag')).toBe(false);
  });

  it('reports a tag it could not push as a failure, without undoing the publish', () => {
    const f = fake({ gitFails: 'push' });
    const report = runRelease(options({ units: { ...nothing, engine: true } }), f.deps);
    expect(report.failures.join('\n')).toContain('engine-v0.1.0');
    expect(report.entries[0].outcome).toBe('published');
  });

  it('does not tag without --tag', () => {
    const f = fake();
    expect(runRelease(options({ units: { ...nothing, engine: true }, tag: false }), f.deps).tags).toEqual([]);
  });
});
