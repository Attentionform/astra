import * as fs from 'node:fs';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';

const root = path.resolve(__dirname, '..', '..');
const read = (rel: string): string => fs.readFileSync(path.join(root, rel), 'utf8');
const workflows = ['ci.yml', 'release.yml', 'validator-platform.yml'].map((f) => `.github/workflows/${f}`);

/** "10.0.403" -> [10, 0, 403] */
function sdkParts(v: string): number[] {
  return v.split('.').map((n) => Number(n));
}

describe('validator native build configuration', () => {
  const pins = workflows.flatMap((wf) =>
    [...read(wf).matchAll(/dotnet-version:\s*'?([^'\s]+)'?/g)].map((m) => ({ wf, version: m[1] }))
  );

  it('every workflow installs the .NET SDK', () => {
    for (const wf of workflows) expect(pins.some((p) => p.wf === wf)).toBe(true);
  });

  it('every workflow pins the same exact SDK version', () => {
    for (const p of pins) expect(p.version).toMatch(/^\d+\.\d+\.\d+$/);
    expect(new Set(pins.map((p) => p.version)).size).toBe(1);
  });

  it('the CI pin satisfies the minimum SDK in global.json', () => {
    const sdk = JSON.parse(read('packages/ooxml-validator/native/global.json')).sdk;
    const min = sdkParts(sdk.version);
    const pinned = sdkParts(pins[0].version);
    expect(sdk.rollForward).toBe('latestFeature');
    expect(pinned[0]).toBe(min[0]);
    expect(pinned[1]).toBe(min[1]);
    expect(pinned[2]).toBeGreaterThanOrEqual(min[2]);
  });

  it('fails the restore on high and critical NuGet advisories', () => {
    const csproj = read('packages/ooxml-validator/native/ooxml-validator.csproj');
    expect(csproj).toMatch(/<NuGetAudit>true<\/NuGetAudit>/);
    expect(csproj).toMatch(/<NuGetAuditMode>all<\/NuGetAuditMode>/);
    expect(csproj).toMatch(/<NuGetAuditLevel>high<\/NuGetAuditLevel>/);
    expect(csproj).toMatch(/WarningsAsErrors[^>]*>[^<]*NU1903;NU1904/);
    // An unreachable advisory feed must not pass silently on CI.
    expect(csproj).toMatch(/Condition="'\$\(CI\)' == 'true'"[\s\S]*NU1900/);
  });

  it('builds all six platforms, each on its own operating system, and never uses the retired macos-13 image', () => {
    const wf = read('.github/workflows/validator-platform.yml');
    const runners: Record<string, string> = {
      'linux-x64': 'ubuntu-latest',
      'linux-arm64': 'ubuntu-24.04-arm',
      'osx-x64': 'macos-15-intel',
      'osx-arm64': 'macos-latest',
      'win-x64': 'windows-latest',
      'win-arm64': 'windows-11-arm',
    };
    for (const [rid, runner] of Object.entries(runners)) {
      expect(wf).toMatch(new RegExp(`rid: ${rid}\\n\\s+npm: [a-z0-9-]+\\n\\s+runner: ${runner}\\n`));
    }
    for (const w of workflows) expect(read(w)).not.toMatch(/macos-13/);
    // Nothing is cross-compiled: the job builds for the machine it runs on.
    expect(wf).not.toMatch(/-r\s+\$\{\{\s*matrix\.rid/);
  });

  it('build jobs never see the npm token: only the platform packages travel, as artifacts, and only when asked to pack', () => {
    const platform = read('.github/workflows/validator-platform.yml');
    expect(platform).not.toMatch(/NPM_TOKEN|NODE_AUTH_TOKEN|secrets\./);
    expect(platform).toMatch(/if: inputs\.pack\n\s+uses: actions\/upload-artifact/);
    expect(read('.github/workflows/validator-check.yml')).not.toMatch(/upload-artifact|download-artifact/);
  });

  it('only the platform workflow caller may ask for the token that provenance needs, and only the publish job does', () => {
    expect(read('.github/workflows/validator-platform.yml')).not.toMatch(/id-token/);
    const release = read('.github/workflows/release.yml');
    expect(release.match(/id-token: write/g)).toHaveLength(1);
    expect(release.slice(release.indexOf('\n  publish:'))).toMatch(/id-token: write/);
  });

  it('the release workflow is manual only: no push, tag or schedule can start it', () => {
    const release = read('.github/workflows/release.yml');
    const on = release.slice(release.indexOf('\non:'), release.indexOf('\nconcurrency:'));
    expect(on).toMatch(/workflow_dispatch:/);
    expect(on).not.toMatch(/\bpush:|\bpull_request:|\bschedule:|\btags:/);
    expect(fs.existsSync(path.join(root, '.github/workflows/publish.yml'))).toBe(false);
  });

  it('a release is a dry run unless someone turns that off, and a real one waits for approval in the npm environment', () => {
    const release = read('.github/workflows/release.yml');
    expect(release).toMatch(/dry_run:\n\s+description:[^\n]*\n\s+type: boolean\n\s+default: true/);
    expect(release).toMatch(/environment: \$\{\{ inputs\.dry_run && 'npm-dry-run' \|\| 'npm' \}\}/);
  });

  it('the npm token is used by the publish job only', () => {
    const release = read('.github/workflows/release.yml');
    const publishStart = release.indexOf('\n  publish:');
    expect(publishStart).toBeGreaterThan(0);
    expect(release.slice(0, publishStart)).not.toMatch(/secrets\.NPM_TOKEN/);
    expect(release.slice(publishStart)).toMatch(/secrets\.NPM_TOKEN/);
  });

  it('lets the publish job run when a platform job failed, but not when the plan or the tests failed', () => {
    const wf = read('.github/workflows/release.yml');
    expect(wf).toMatch(/needs: \[plan, test, platforms\]/);
    expect(wf).toMatch(/!cancelled\(\)/);
    expect(wf).toMatch(/needs\.plan\.result == 'success'/);
    expect(wf).toMatch(/needs\.test\.result == 'success' \|\| needs\.test\.result == 'skipped'/);
    expect(wf).toMatch(/needs\.platforms\.result == 'failure'/);
  });

  it('tests on the lowest supported Node and the newest', () => {
    const wf = read('.github/workflows/release.yml');
    const floor = JSON.parse(read('package.json')).engines.node.replace('>=', '');
    expect(wf).toContain(`node-version: ['${floor}', 'latest']`);
  });
});
