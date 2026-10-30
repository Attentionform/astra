// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));

function detectRid(): string {
  const platform = process.platform;
  const arch = process.arch;

  if (platform === 'darwin' && arch === 'arm64') return 'darwin-arm64';
  if (platform === 'darwin' && arch === 'x64') return 'darwin-x64';
  if (platform === 'linux' && arch === 'arm64') return 'linux-arm64';
  if (platform === 'linux' && arch === 'x64') return 'linux-x64';
  if (platform === 'win32' && arch === 'arm64') return 'win32-arm64';
  if (platform === 'win32' && arch === 'x64') return 'win32-x64';

  throw new Error(`No embedded OOXML validator binary for platform: ${platform} ${arch}`);
}

function dotnetRid(npmRid: string): string {
  if (npmRid.startsWith('darwin-')) return npmRid.replace('darwin-', 'osx-');
  if (npmRid.startsWith('win32-')) return npmRid.replace('win32-', 'win-');
  return npmRid;
}

export function resolveBinaryDetails(): { binaryPath: string; platformPackage: string } {
  // 1. Explicit environment variable override
  const env = process.env.OOXML_VALIDATOR_CLI;
  if (env && env.trim().length > 0) {
    return {
      binaryPath: env.trim(),
      platformPackage: 'env:OOXML_VALIDATOR_CLI',
    };
  }

  const npmRid = detectRid();
  const binName = npmRid.startsWith('win32-') ? 'ooxml-validator.exe' : 'ooxml-validator';
  const pkgName = `@attentionform/ooxml-validator-${npmRid}`;

  // 2. Platform optionalDependency package
  let requireErr: unknown;
  try {
    const resolved = require.resolve(`${pkgName}/${binName}`);
    return {
      binaryPath: resolved,
      platformPackage: pkgName,
    };
  } catch (e) {
    requireErr = e;
    // 3. Local development fallback (built by dotnet publish into native/out/<rid>/)
    const localRid = dotnetRid(npmRid);
    const localDevPath = path.resolve(__dirname, '..', 'native', 'out', localRid, binName);
    if (fs.existsSync(localDevPath)) {
      return {
        binaryPath: localDevPath,
        platformPackage: 'local:native/out',
      };
    }

    throw new Error(
      `Missing optional dependency ${pkgName} for ${process.platform}-${process.arch} (${pkgName}/${binName}). ` +
        `Local dev binary not found at ${localDevPath}. ` +
        'Try building locally (npm run build:native) or set OOXML_VALIDATOR_CLI to your own CLI path.' +
        (requireErr ? ` Underlying error: ${(requireErr as Error).message}` : ''),
    );
  }
}

export function resolveEmbeddedBinary(): string {
  return resolveBinaryDetails().binaryPath;
}
