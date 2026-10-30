export interface ValidatorTarget {
  rid: string;          // .NET Runtime Identifier
  npmSuffix: string;    // npm package suffix, e.g. "linux-x64"
  os: 'linux' | 'darwin' | 'win32';
  cpu: 'x64' | 'arm64';
  binaryName: string;
}

export const VALIDATOR_TARGETS: readonly ValidatorTarget[] = [
  { rid: 'linux-x64', npmSuffix: 'linux-x64', os: 'linux', cpu: 'x64', binaryName: 'ooxml-validator' },
  { rid: 'linux-arm64', npmSuffix: 'linux-arm64', os: 'linux', cpu: 'arm64', binaryName: 'ooxml-validator' },
  { rid: 'osx-x64', npmSuffix: 'darwin-x64', os: 'darwin', cpu: 'x64', binaryName: 'ooxml-validator' },
  { rid: 'osx-arm64', npmSuffix: 'darwin-arm64', os: 'darwin', cpu: 'arm64', binaryName: 'ooxml-validator' },
  { rid: 'win-x64', npmSuffix: 'win32-x64', os: 'win32', cpu: 'x64', binaryName: 'ooxml-validator.exe' },
  { rid: 'win-arm64', npmSuffix: 'win32-arm64', os: 'win32', cpu: 'arm64', binaryName: 'ooxml-validator.exe' },
];

export function getHostTarget(): ValidatorTarget {
  const platform = process.platform;
  const arch = process.arch;
  const target = VALIDATOR_TARGETS.find(t => t.os === platform && t.cpu === arch);
  if (!target) {
    throw new Error(`Unsupported host platform/architecture: ${platform}-${arch}`);
  }
  return target;
}
