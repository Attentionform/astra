# @attentionform/ooxml-validator

Cross-platform OOXML (ECMA-376 / ISO-29500) document schema validator powered by .NET 10 OpenXmlValidator.

Portions of the code are adapted from `xarsh/ooxml-validator` (MIT). See `THIRD_PARTY_NOTICES.md`.

## Installation

```bash
npm install @attentionform/ooxml-validator
```

The package automatically uses the matching prebuilt native platform binary from optional dependencies:
- `@attentionform/ooxml-validator-linux-x64`
- `@attentionform/ooxml-validator-linux-arm64`
- `@attentionform/ooxml-validator-darwin-x64`
- `@attentionform/ooxml-validator-darwin-arm64`
- `@attentionform/ooxml-validator-win32-x64`
- `@attentionform/ooxml-validator-win32-arm64`

## API

### `validateFile(filePath, options?): Promise<ValidationResult>`
Validates an OOXML document from disk. Defaults to `Microsoft365` office version when unspecified.

```typescript
import { validateFile } from '@attentionform/ooxml-validator';

const result = await validateFile('path/to/doc.docx', { officeVersion: 'Office2019' });
console.log(result.ok, result.errors);
```

### `isValid(filePath, options?): Promise<boolean>`
Convenience helper returning `true` when the file validates without schema errors.

```typescript
import { isValid } from '@attentionform/ooxml-validator';

if (!await isValid('contract.docx')) {
  throw new Error('Invalid document');
}
```

### `validateBuffer(bytes, options?): Promise<ValidationResult>`
Validates an in-memory document buffer directly without creating temporary disk files. Uses a resident worker process via streaming stdin/stdout IPC. Defaults to `Office2019`.

```typescript
import { validateBuffer } from '@attentionform/ooxml-validator';
import fs from 'node:fs';

const bytes = fs.readFileSync('document.docx');
const result = await validateBuffer(bytes, { officeVersion: 'Office2019', maxErrors: 100 });
```

### `createValidator(options?): ValidatorInstance`
Creates an isolated validator worker client with configurable recycling, limits, and timeouts.

```typescript
import { createValidator } from '@attentionform/ooxml-validator';

const validator = createValidator({
  timeoutMs: 20000,
  recycleAfter: 500,
  maxDocBytes: 64 * 1024 * 1024,
});

const result = await validator.validate(buffer);
await validator.close();
```

### `selfTest(): Promise<{ ok: boolean; details: string }>`
Validates internal known-good and known-bad embedded test documents to verify runtime health.

### `backendInfo(): BackendInfo`
Returns diagnostic metadata regarding the active binary, platform package, SDK version, and SHA-256 hash.

## Environment Variables

- `OOXML_VALIDATOR_CLI`: Explicit path to a native `ooxml-validator` binary. Overrides platform package resolution and local fallback.

## Supported Platforms

- Linux (x64, arm64) - glibc 2.31+ (Linux musl / Alpine is currently not supported)
- macOS (Apple Silicon arm64, Intel x64) - macOS 13+
- Windows (x64, arm64) - Windows 10+ / Windows Server 2019+

## Local Development & Building Native Binaries

Building the native executable requires the .NET 10 SDK, version 10.0.100 or newer (the minimum in `native/global.json`). CI builds with the exact version set as `dotnet-version` in `.github/workflows/`; all three workflows must carry the same version (a test checks this). To bump it, change the version in all three workflows, then regenerate the lock file with that SDK (`dotnet restore native/ooxml-validator.csproj --force-evaluate`): the lock records runtime-tied packages, so a locked restore fails on a different SDK until it is regenerated. The restore also audits the dependencies and fails on a known high or critical advisory.

```bash
# Build the native binary for the current host architecture:
npm run build:native

# Compile the TypeScript wrapper:
npm run build

# Run unit tests:
npm test
```

## License

This package is licensed under BUSL-1.1. It incorporates and adapts code from `xarsh/ooxml-validator` (Copyright (c) 2025 Takao Naito, MIT License) and uses Microsoft's DocumentFormat.OpenXml SDK (MIT License). See `THIRD_PARTY_NOTICES.md` for full attribution and license texts.

The six platform packages (`@attentionform/ooxml-validator-<platform>`) are not listed in this package's `package.json` in the repository: they are built from the native binaries at release time, and listing packages that are not on npm yet would make `npm ci` fail. The published wrapper lists them as optional dependencies at the range of native binaries it accepts, which is declared in the repository manifest as `validatorBinary.range` (for example `>=0.1.0 <0.2.0`). The platform packages carry that binary version, which changes only when the native validator changes and is separate from the wrapper's own version. So a platform whose newest binary could not be published keeps the last one inside the range, and the wrapper is not released while any platform has no binary in the range. See [Releasing](../../docs/releasing.md).
