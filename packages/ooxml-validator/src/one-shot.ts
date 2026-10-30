// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import { spawn } from 'node:child_process';
import { resolveEmbeddedBinary } from './resolve-binary.js';
import { OFFICE_VERSIONS, type OfficeVersion, type ValidateOptions, type ValidationResult } from './types.js';

function isOfficeVersion(value: string): value is OfficeVersion {
  return (OFFICE_VERSIONS as readonly string[]).includes(value);
}

export function runValidatorOneShot(file: string, options: ValidateOptions = {}): Promise<ValidationResult> {
  const cliArgs = [file];

  if (options.officeVersion !== undefined) {
    if (!isOfficeVersion(options.officeVersion)) {
      return Promise.reject(
        new TypeError(
          `Invalid officeVersion ${JSON.stringify(options.officeVersion)}. Expected one of: ${OFFICE_VERSIONS.join(', ')}`,
        ),
      );
    }
    cliArgs.push(options.officeVersion);
  }

  let cmd: string;
  try {
    cmd = resolveEmbeddedBinary();
  } catch (e) {
    return Promise.reject(e as Error);
  }

  return new Promise((resolve, reject) => {
    const child = spawn(cmd, cliArgs, { stdio: ['ignore', 'pipe', 'pipe'] });

    let stdout = '';
    let stderr = '';

    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');

    child.stdout.on('data', (d) => {
      stdout += d;
    });
    child.stderr.on('data', (d) => {
      stderr += d;
    });

    child.on('error', (err) => {
      reject(new Error(`Failed to spawn OOXML validator: ${err.message}`));
    });

    child.on('close', (code) => {
      if (code !== 0) {
        return reject(new Error(`OOXML Validator exited with code ${code}. stderr: ${stderr || stdout}`));
      }

      const trimmed = stdout.trim();
      if (!trimmed) {
        return reject(new Error('OOXML Validator produced no output.'));
      }

      try {
        resolve(JSON.parse(trimmed) as ValidationResult);
      } catch (e) {
        reject(new Error(`Failed to parse OOXML validator output: ${(e as Error).message}\nRaw output: ${trimmed}`));
      }
    });
  });
}
