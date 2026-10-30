// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import fs from 'node:fs';
import crypto from 'node:crypto';
import { runValidatorOneShot } from './one-shot.js';
import { resolveBinaryDetails } from './resolve-binary.js';
export { resolveEmbeddedBinary, resolveBinaryDetails } from './resolve-binary.js';
import { WorkerClient } from './worker-client.js';
import {
  OFFICE_VERSIONS,
  type OfficeVersion,
  type ValidateOptions,
  type ValidationError,
  type ValidationResult,
  type ValidatorPoolOptions,
  type BackendInfo,
  type ValidatorInstance,
  type ErrorNode,
  type ValidationErrorClass,
} from './types.js';

export { OFFICE_VERSIONS };
export type {
  OfficeVersion,
  ValidateOptions,
  ValidationError,
  ValidationResult,
  ValidatorPoolOptions,
  BackendInfo,
  ValidatorInstance,
  ErrorNode,
  ValidationErrorClass,
};

let defaultWorker: WorkerClient | null = null;
let cachedSha256: string | null = null;

function getDefaultWorker(): WorkerClient {
  if (!defaultWorker) {
    defaultWorker = new WorkerClient();
    process.once('exit', () => {
      defaultWorker?.close().catch(() => {});
    });
  }
  return defaultWorker;
}

/**
 * Validates a single OOXML file from disk.
 * Defaults to 'Microsoft365' to match standard ECMA-376 / ISO-29500 behavior.
 */
export async function validateFile(
  filePath: string,
  options: ValidateOptions = {},
): Promise<ValidationResult> {
  return runValidatorOneShot(filePath, options);
}

/**
 * Convenience helper returning a boolean indicating if the document is valid.
 */
export async function isValid(
  filePath: string,
  options: ValidateOptions = {},
): Promise<boolean> {
  const res = await validateFile(filePath, options);
  return res.ok;
}

/**
 * Validates an in-memory DOCX buffer without creating temporary files.
 * Defaults to 'Office2019' matching the ASTra engine's standard validation target.
 */
export async function validateBuffer(
  bytes: Buffer | Uint8Array,
  options: ValidateOptions = {},
): Promise<ValidationResult> {
  const worker = getDefaultWorker();
  const officeVersion = options.officeVersion ?? 'Office2019';
  const size = bytes.length;

  const resp = await worker.send(
    {
      op: 'validate',
      kind: 'docx',
      office: officeVersion,
      size,
      maxErrors: options.maxErrors,
      maxDocBytes: options.maxDocBytes,
    },
    bytes,
    options.timeoutMs,
  );

  return {
    file: '<memory>',
    ok: resp.ok,
    sdkVersion: resp.sdkVersion,
    office: resp.office ?? officeVersion,
    durationMs: resp.durationMs,
    errorCount: resp.errorCount ?? (resp.errors ? resp.errors.length : 0),
    truncated: resp.truncated ?? false,
    errors: resp.errors ?? [],
  };
}

/**
 * Creates an isolated validator worker client instance.
 */
export function createValidator(opts: ValidatorPoolOptions = {}): ValidatorInstance {
  const worker = new WorkerClient(opts);

  return {
    async validate(bytes: Buffer | Uint8Array, options: ValidateOptions = {}): Promise<ValidationResult> {
      const officeVersion = options.officeVersion ?? 'Office2019';
      const resp = await worker.send(
        {
          op: 'validate',
          kind: 'docx',
          office: officeVersion,
          size: bytes.length,
          maxErrors: options.maxErrors,
          maxDocBytes: options.maxDocBytes,
        },
        bytes,
        options.timeoutMs,
      );

      return {
        file: '<memory>',
        ok: resp.ok,
        sdkVersion: resp.sdkVersion,
        office: resp.office ?? officeVersion,
        durationMs: resp.durationMs,
        errorCount: resp.errorCount ?? (resp.errors ? resp.errors.length : 0),
        truncated: resp.truncated ?? false,
        errors: resp.errors ?? [],
      };
    },

    async validateFile(filePath: string, options: ValidateOptions = {}): Promise<ValidationResult> {
      const bytes = await fs.promises.readFile(filePath);
      const res = await this.validate(bytes, options);
      return { ...res, file: filePath };
    },

    async close(): Promise<void> {
      await worker.close();
    },
  };
}

let cachedSdkVersion: string | null = null;

function computeFileSha256(filePath: string): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const hash = crypto.createHash('sha256');
    const stream = fs.createReadStream(filePath);
    stream.on('data', (chunk) => hash.update(chunk));
    stream.on('end', () => resolve(hash.digest('hex')));
    stream.on('error', (err) => reject(err));
  });
}

/**
 * Returns diagnostic metadata about the resolved binary and platform package.
 */
export async function backendInfo(): Promise<BackendInfo> {
  const { binaryPath, platformPackage } = resolveBinaryDetails();
  if (!cachedSha256) {
    try {
      cachedSha256 = await computeFileSha256(binaryPath);
    } catch {
      cachedSha256 = 'unknown';
    }
  }

  if (!cachedSdkVersion) {
    try {
      const worker = getDefaultWorker();
      const res = await worker.send({ op: 'version' });
      if (res.ok && res.sdkVersion) {
        cachedSdkVersion = res.sdkVersion;
      }
    } catch {
      // fallback
    }
  }

  return {
    platformPackage,
    binaryPath,
    sdkVersion: cachedSdkVersion ?? '3.5.1',
    binarySha256: cachedSha256,
    workerSupported: true,
  };
}

// Embedded minimal valid DOCX base64 for zero-dependency self-test
const MINIMAL_VALID_DOCX_BASE64 =
  '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';

// Embedded invalid schema DOCX base64 (from tests/fixtures/invalid-schema.docx)
const INVALID_SCHEMA_DOCX_BASE64 =
  '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';

/**
 * Runs a self-diagnostic validating an embedded known-good and known-bad document.
 */
export async function selfTest(): Promise<{ ok: boolean; details: string }> {
  try {
    const validBytes = Buffer.from(MINIMAL_VALID_DOCX_BASE64, 'base64');
    const invalidBytes = Buffer.from(INVALID_SCHEMA_DOCX_BASE64, 'base64');

    const validRes = await validateBuffer(validBytes);
    const invalidRes = await validateBuffer(invalidBytes);

    if (!validRes.ok) {
      return {
        ok: false,
        details: `Known-valid test document failed validation: ${JSON.stringify(validRes.errors)}`,
      };
    }

    if (invalidRes.ok) {
      return {
        ok: false,
        details: 'Known-invalid test document unexpectedly passed validation',
      };
    }

    const firstError = invalidRes.errors[0];
    if (!firstError || firstError.class !== 'schema') {
      return {
        ok: false,
        details: `Known-invalid test document did not produce expected schema error class (got: ${firstError?.class})`,
      };
    }

    const info = await backendInfo();
    return {
      ok: true,
      details: `self-test passed: valid and invalid documents verified against .NET 10 OpenXmlValidator ${info.sdkVersion} (${info.platformPackage})`,
    };
  } catch (err) {
    return {
      ok: false,
      details: `self-test threw: ${(err as Error).message}`,
    };
  }
}
