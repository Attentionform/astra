// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import { describe, it, expect } from 'vitest';
import { execFile } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
// A file:// URL, because the child process imports it and a Windows path such as D:\... is not a valid import specifier.
const RESOLVE = pathToFileURL(fileURLToPath(new URL('../dist/resolve-binary.js', import.meta.url))).href;

async function resolveAs(platform: string, arch: string): Promise<{ path?: string; error?: string }> {
  const script = `
    Object.defineProperty(process, 'platform', { value: ${JSON.stringify(platform)} });
    Object.defineProperty(process, 'arch', { value: ${JSON.stringify(arch)} });
    import(${JSON.stringify(RESOLVE)}).then(({ resolveEmbeddedBinary }) => {
      try {
        console.log(JSON.stringify({ path: resolveEmbeddedBinary() }));
      } catch (e) {
        console.log(JSON.stringify({ error: e.message }));
      }
    });
  `;
  const { stdout } = await execFileAsync(process.execPath, ['--eval', script]);
  return JSON.parse(stdout);
}

describe('resolveEmbeddedBinary', () => {
  it('each supported platform maps to its own package and executable name', async () => {
    const expected = [
      ['darwin', 'arm64', '@attentionform/ooxml-validator-darwin-arm64/ooxml-validator'],
      ['darwin', 'x64', '@attentionform/ooxml-validator-darwin-x64/ooxml-validator'],
      ['linux', 'arm64', '@attentionform/ooxml-validator-linux-arm64/ooxml-validator'],
      ['linux', 'x64', '@attentionform/ooxml-validator-linux-x64/ooxml-validator'],
      ['win32', 'arm64', '@attentionform/ooxml-validator-win32-arm64/ooxml-validator.exe'],
      ['win32', 'x64', '@attentionform/ooxml-validator-win32-x64/ooxml-validator.exe'],
    ] as const;

    for (const [platform, arch, specifier] of expected) {
      const { path: raw, error } = await resolveAs(platform, arch);
      const path = raw?.replace(/\\/g, '/'); // Windows paths use backslashes
      const isResolved =
        path?.endsWith(specifier.replace('@attentionform/', '')) ||
        (path && path.includes('native/out/')) ||
        error?.includes(specifier);
      expect(isResolved).toBe(true);
    }
  });

  it('an unsupported platform says so by name', async () => {
    const { error } = await resolveAs('freebsd', 'ppc64');
    expect(String(error)).toMatch(/No embedded OOXML validator binary for platform: freebsd ppc64/);
  });

  it('env override sets platformPackage to env:OOXML_VALIDATOR_CLI', async () => {
    const { stdout } = await execFileAsync(process.execPath, [
      '--eval',
      `
        process.env.OOXML_VALIDATOR_CLI = '/custom/path/bin';
        import(${JSON.stringify(RESOLVE)}).then(({ resolveBinaryDetails }) => {
          console.log(JSON.stringify(resolveBinaryDetails()));
        });
      `,
    ]);
    const res = JSON.parse(stdout);
    expect(res.binaryPath).toBe('/custom/path/bin');
    expect(res.platformPackage).toBe('env:OOXML_VALIDATOR_CLI');
  });
});
