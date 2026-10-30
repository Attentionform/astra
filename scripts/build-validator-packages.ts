#!/usr/bin/env tsx
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { VALIDATOR_TARGETS } from './validator-targets.js';
import { readValidatorBinary } from './validator-binary.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, '..');
const wrapperPkgPath = path.join(repoRoot, 'packages', 'ooxml-validator', 'package.json');
const defaultOutDir = path.join(repoRoot, 'packages', 'ooxml-validator', 'dist-platform');

/** Copied next to each published binary by the csproj; shipped in every platform package. */
export const RUNTIME_NOTICE_FILES = ['THIRD-PARTY-NOTICES.TXT', 'DOTNET-LICENSE.TXT'];

function parseArgs(): { binariesDir: string; onlyTarget?: string; outDir: string } {
  const args = process.argv.slice(2);
  let binariesDir = path.join(repoRoot, 'packages', 'ooxml-validator', 'native', 'out');
  let onlyTarget: string | undefined;
  let outDir = defaultOutDir;

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--binaries' && i + 1 < args.length) {
      binariesDir = path.resolve(args[++i]);
    } else if (args[i] === '--only' && i + 1 < args.length) {
      onlyTarget = args[++i];
    } else if (args[i] === '--out-dir' && i + 1 < args.length) {
      outDir = path.resolve(args[++i]);
    }
  }

  return { binariesDir, onlyTarget, outDir };
}

function computeSha256(filePath: string): string {
  const data = fs.readFileSync(filePath);
  return crypto.createHash('sha256').update(data).digest('hex');
}

export function buildValidatorPackages({
  binariesDir,
  onlyTarget,
  outDir,
}: {
  binariesDir: string;
  onlyTarget?: string;
  outDir: string;
}): { builtTargets: string[]; checksumsFile: string } {
  if (!fs.existsSync(wrapperPkgPath)) {
    throw new Error(`Wrapper package.json not found at: ${wrapperPkgPath}`);
  }
  // The platform packages carry the binary version, not the wrapper's (see validator-binary.ts).
  const version = readValidatorBinary(wrapperPkgPath).version;

  const licenseContent = fs.readFileSync(path.join(repoRoot, 'LICENSE'), 'utf8');
  const noticeContent = fs.readFileSync(path.join(repoRoot, 'NOTICE'), 'utf8');
  const tpNoticesContent = fs.readFileSync(path.join(repoRoot, 'THIRD_PARTY_NOTICES.md'), 'utf8');

  fs.mkdirSync(outDir, { recursive: true });

  const builtTargets: string[] = [];
  const checksumLines: string[] = [];

  for (const target of VALIDATOR_TARGETS) {
    if (onlyTarget && target.rid !== onlyTarget && target.npmSuffix !== onlyTarget) {
      continue;
    }

    // Binary layout: <binariesDir>/<rid>/ooxml-validator[.exe]
    const binPath = path.join(binariesDir, target.rid, target.binaryName);
    if (!fs.existsSync(binPath)) {
      if (onlyTarget) {
        throw new Error(`Expected binary not found for target ${target.rid} at ${binPath}`);
      }
      // Target binary not present in binariesDir; skip when scanning all targets
      continue;
    }

    const pkgDirName = target.npmSuffix;
    const pkgDir = path.join(outDir, pkgDirName);
    fs.mkdirSync(pkgDir, { recursive: true });

    // 1. Copy binary and set executable permissions
    const destBin = path.join(pkgDir, target.binaryName);
    fs.copyFileSync(binPath, destBin);
    if (target.os !== 'win32') {
      fs.chmodSync(destBin, 0o755);
    }

    const sha256 = computeSha256(destBin);
    checksumLines.push(`${sha256}  ${target.npmSuffix}/${target.binaryName}`);

    // 2. Copy legal notices from root
    fs.writeFileSync(path.join(pkgDir, 'LICENSE'), licenseContent, 'utf8');
    fs.writeFileSync(path.join(pkgDir, 'NOTICE'), noticeContent, 'utf8');
    fs.writeFileSync(path.join(pkgDir, 'THIRD_PARTY_NOTICES.md'), tpNoticesContent, 'utf8');

    // The binary embeds the .NET runtime; its licence and third-party notices (copied next to the binary by
    // the csproj publish target) must ship with it. A package without them is not built.
    const extraFiles = RUNTIME_NOTICE_FILES;
    for (const name of extraFiles) {
      const src = path.join(binariesDir, target.rid, name);
      if (!fs.existsSync(src)) {
        throw new Error(
          `Missing ${name} for target ${target.rid} at ${src}. It is copied from the .NET runtime pack when the binary is published (see the CopyRuntimeNotices target in ooxml-validator.csproj).`
        );
      }
      fs.copyFileSync(src, path.join(pkgDir, name));
    }

    // 3. Write README.md and SHA256SUMS
    const readmeContent = [
      `# @attentionform/ooxml-validator-${target.npmSuffix}`,
      '',
      `Prebuilt native OOXML validator binary for ${target.os}-${target.cpu} (\`${target.rid}\`).`,
      '',
      `- Binary: \`${target.binaryName}\``,
      `- SHA-256: \`${sha256}\``,
      '',
      'This package is an optional platform dependency of `@attentionform/ooxml-validator`.',
    ].join('\n');
    fs.writeFileSync(path.join(pkgDir, 'README.md'), readmeContent, 'utf8');
    fs.writeFileSync(path.join(pkgDir, 'SHA256SUMS'), `${sha256}  ${target.binaryName}\n`, 'utf8');

    // 4. Generate package.json
    const files = [
      target.binaryName,
      'LICENSE',
      'NOTICE',
      'THIRD_PARTY_NOTICES.md',
      'README.md',
      'SHA256SUMS',
      ...extraFiles,
    ];

    const pkgJson = {
      name: `@attentionform/ooxml-validator-${target.npmSuffix}`,
      version,
      description: `Prebuilt native OOXML validator binary for ${target.os}-${target.cpu}`,
      os: [target.os],
      cpu: [target.cpu],
      main: target.binaryName,
      files,
      license: 'BUSL-1.1',
      repository: {
        type: 'git',
        url: 'git+https://github.com/Attentionform/astra.git',
      },
      publishConfig: {
        access: 'public',
      },
    };

    fs.writeFileSync(path.join(pkgDir, 'package.json'), JSON.stringify(pkgJson, null, 2) + '\n', 'utf8');
    builtTargets.push(target.npmSuffix);
    console.log(`[build-validator-packages] Built ${target.npmSuffix} (${sha256.slice(0, 16)}...)`);
  }

  // 5. Write checksums.txt for all built targets
  const checksumsFile = path.join(outDir, 'checksums.txt');
  fs.writeFileSync(checksumsFile, checksumLines.join('\n') + (checksumLines.length ? '\n' : ''), 'utf8');
  console.log(`[build-validator-packages] Wrote ${builtTargets.length} package(s) to ${outDir}`);

  return { builtTargets, checksumsFile };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const options = parseArgs();
  buildValidatorPackages(options);
}
