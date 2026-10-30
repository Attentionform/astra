import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, it, expect } from 'vitest';
import { isPublished, publishIfNew, publishedVersions, type NpmResult, type NpmRunner } from '../../scripts/npm-publish-if-new.js';
import { judgePlatforms, platformState } from '../../scripts/validator-platform-status.js';

const ok = (stdout = ''): NpmResult => ({ status: 0, stdout, stderr: '' });
const notFound: NpmResult = { status: 1, stdout: '', stderr: 'npm error code E404' };

const dirs: string[] = [];
function packageDir(name = '@scope/pkg', version = '0.1.0'): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-publish-'));
  dirs.push(dir);
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name, version }));
  return dir;
}
afterEach(() => {
  for (const d of dirs.splice(0)) fs.rmSync(d, { recursive: true, force: true });
});

/** A fake npm: `view` answers from `published` (which `publish` extends), every call is recorded. */
function fakeNpm(published: Set<string>, calls: string[][]): NpmRunner {
  return (args) => {
    calls.push(args);
    if (args[0] === 'view') {
      const spec = args[1];
      return published.has(spec) ? ok(spec.slice(spec.lastIndexOf('@') + 1) + '\n') : notFound;
    }
    if (args[0] === 'publish' && !args.includes('--dry-run')) published.add('@scope/pkg@0.1.0');
    return ok();
  };
}

describe('isPublished', () => {
  it('is true when the registry serves the version and false on E404', () => {
    expect(isPublished('a', '1.0.0', () => ok('1.0.0\n'))).toBe(true);
    expect(isPublished('a', '1.0.0', () => notFound)).toBe(false);
  });
  it('throws when the registry cannot be asked, instead of guessing', () => {
    expect(() => isPublished('a', '1.0.0', () => ({ status: 1, stdout: '', stderr: 'npm error network ETIMEDOUT' }))).toThrow(/Could not ask the registry/);
  });
});

describe('publishIfNew', () => {
  it('skips a version that is already on npm and does not publish', () => {
    const calls: string[][] = [];
    const out = publishIfNew({ dir: packageDir(), run: fakeNpm(new Set(['@scope/pkg@0.1.0']), calls), wait: () => {} });
    expect(out).toBe('skipped');
    expect(calls.some((c) => c[0] === 'publish')).toBe(false);
  });

  it('publishes a new version publicly and checks that the registry serves it', () => {
    const calls: string[][] = [];
    const dir = packageDir();
    const out = publishIfNew({ dir, run: fakeNpm(new Set(), calls), wait: () => {} });
    expect(out).toBe('published');
    const publish = calls.find((c) => c[0] === 'publish')!;
    expect(publish).toEqual(['publish', dir, '--access', 'public']);
    expect(calls[calls.length - 1][0]).toBe('view');
  });

  it('passes --provenance only when asked', () => {
    const calls: string[][] = [];
    publishIfNew({ dir: packageDir(), provenance: true, run: fakeNpm(new Set(), calls), wait: () => {} });
    expect(calls.find((c) => c[0] === 'publish')).toContain('--provenance');
  });

  it('does a dry run without asking the registry or verifying', () => {
    const calls: string[][] = [];
    const out = publishIfNew({ dir: packageDir(), dryRun: true, run: fakeNpm(new Set(['@scope/pkg@0.1.0']), calls), wait: () => {} });
    expect(out).toBe('dry-run');
    expect(calls).toHaveLength(1);
    expect(calls[0]).toContain('--dry-run');
  });

  it('fails when npm publish fails', () => {
    const run: NpmRunner = (args) => (args[0] === 'view' ? notFound : { status: 1, stdout: '', stderr: 'npm error E403' });
    expect(() => publishIfNew({ dir: packageDir(), run, wait: () => {} })).toThrow(/npm publish failed/);
  });

  it('fails when the registry never serves the new version', () => {
    const run: NpmRunner = (args) => (args[0] === 'view' ? notFound : ok());
    let waits = 0;
    expect(() => publishIfNew({ dir: packageDir(), run, attempts: 3, wait: () => waits++ })).toThrow(/does not serve it yet/);
    expect(waits).toBe(2);
  });
});

describe('publishedVersions', () => {
  it('lists the versions the registry has', () => {
    expect(publishedVersions('@s/p', () => ok('["0.1.0","0.1.1"]'))).toEqual(['0.1.0', '0.1.1']);
  });
  it('reads the single-version answer npm gives as a bare string', () => {
    expect(publishedVersions('@s/p', () => ok('"0.1.0"'))).toEqual(['0.1.0']);
  });
  it('says there are none for a package that does not exist', () => expect(publishedVersions('@s/p', () => notFound)).toEqual([]));
  it('throws when the registry cannot be asked, instead of guessing', () => {
    expect(() => publishedVersions('@s/p', () => ({ status: 1, stdout: '', stderr: 'npm error network timeout' }))).toThrow(/timeout/);
  });
});

describe('judgePlatforms (what the wrapper job does with the registry answer)', () => {
  const binary = { version: '0.2.0', range: '>=0.1.0 <0.3.0' };
  const six = ['linux-x64', 'linux-arm64', 'darwin-x64', 'darwin-arm64', 'win32-x64', 'win32-arm64'];
  const states = (onNpm: Record<string, string[]>) => six.map((s) => platformState(s, binary, onNpm[s] ?? []));
  const all = (versions: string[]) => Object.fromEntries(six.map((s) => [s, versions]));

  it('lets the wrapper go when every platform has the current binary', () => {
    expect(judgePlatforms(states(all(['0.1.0', '0.2.0'])), false)).toMatchObject({ ok: true });
  });
  it('lets the wrapper go when a platform kept an older binary inside the range, and says which', () => {
    const v = judgePlatforms(states({ ...all(['0.2.0']), 'linux-arm64': ['0.1.0', '0.1.4'] }), false);
    expect(v.ok).toBe(true);
    expect(v.message).toContain('linux-arm64 (stays on 0.1.4)');
  });
  it('stops the wrapper when a platform has no binary inside the range', () => {
    const v = judgePlatforms(states({ ...all(['0.2.0']), 'win32-arm64': ['0.0.9', '0.3.0'] }), false);
    expect(v.ok).toBe(false);
    expect(v.message).toContain('win32-arm64');
  });
  it('stops the wrapper when nothing is on npm yet', () => expect(judgePlatforms(states({}), false).ok).toBe(false));
  it('with require-all, stops the wrapper when one platform is only on an older binary', () => {
    const v = judgePlatforms(states({ ...all(['0.2.0']), 'linux-arm64': ['0.1.4'] }), true);
    expect(v.ok).toBe(false);
    expect(v.message).toContain('linux-arm64');
  });
  it('counts a binary being published in this run as current', () => {
    const list = six.map((s) => platformState(s, binary, [], true));
    expect(judgePlatforms(list, true).ok).toBe(true);
  });
  it('does not count a binary outside the range, even a newer one', () => {
    expect(platformState('linux-x64', binary, ['0.3.0']).inRange).toEqual([]);
  });
});
