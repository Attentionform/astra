/**
 * OOXML Validation Gate using @attentionform/ooxml-validator (.NET OpenXmlValidator)
 * 
 * Validates .docx buffers or files against official ECMA-376 / ISO-29500 XSD schemas.
 * Rejects corrupt documents before they are saved to disk or delivered to users.
 */

import { getValidatorBackend, type ValidationResult, type ValidateOptions } from './validator-backend';

export type { ValidationResult, ValidateOptions };

export interface ValidationError {
  description?: string | null;
  path?: string | null;
  xPath?: string | null;
  id?: string | null;
  errorType?: string | null;
  [key: string]: unknown;
}

/**
 * Parses the 1-based body-top-level table ordinal from a validator xPath
 * (e.g. `/w:document[1]/w:body[1]/w:tbl[2]/...` -> 2). Returns null when the
 * xPath names no indexed table. Pure string parse, never throws.
 * Ordinals count top-level body tables only: a nested-table fault xPath
 * (e.g. `.../w:tbl[2]/.../w:tbl[1]/...`) resolves to the outer containing
 * table (first w:tbl index), never a wrong sibling.
 */
export function parseTableOrdinalFromXPath(xPath: unknown): number | null {
  if (typeof xPath !== 'string' || !xPath) return null;
  const m = /w:tbl\[(\d+)\]/i.exec(xPath);
  if (!m) return null;
  const n = Number.parseInt(m[1], 10);
  return Number.isFinite(n) && n >= 1 ? n : null;
}

/**
 * Short human location hint for a validator xPath (`table 2` when the xPath
 * carries an indexed table, `table` when it names a table without an index).
 * Null when the xPath is not table-related. Never throws.
 */
export function tableIndexHintForXPath(xPath: unknown): string | null {
  const n = parseTableOrdinalFromXPath(xPath);
  if (n !== null) return `table ${n}`;
  if (typeof xPath === 'string' && xPath && /w:tbl/i.test(xPath)) return 'table';
  return null;
}

export class OoxmlValidationError extends Error {
  public readonly errors: ValidationError[];
  public readonly firstError: ValidationError;
  public readonly code = 'E_VALIDATION' as const;
  public readonly remedy = 're-export the file from Word (or pick a valid .docx) and retry';
  public readonly retryable = false as const;

  constructor(errors: ValidationError[], context: string = 'Document') {
    const first = errors[0] || {};
    const xp = (typeof first.xPath === 'string' && first.xPath) || (typeof first.path === 'string' && first.path) || 'unknown';
    const hint = tableIndexHintForXPath((first as { xPath?: unknown }).xPath ?? (first as { path?: unknown }).path);
    const detail = `${first.description || 'Schema error'} @ ${xp}${hint ? ` (${hint})` : ''}`;
    super(`[OOXML Schema Violation] ${context} has ${errors.length} validation error(s): ${detail}`);
    this.name = 'OoxmlValidationError';
    this.errors = errors;
    this.firstError = first;
  }
}

/**
 * Validates a DOCX buffer directly in memory using the native validator backend.
 */
export async function validateDocxBuffer(buffer: Buffer, options: ValidateOptions = {}): Promise<ValidationResult> {
  const backend = getValidatorBackend();
  return backend.validateBuffer(buffer, options);
}

/**
 * Validates a DOCX file from disk using the native validator backend.
 */
export async function validateDocxFile(filePath: string, options: ValidateOptions = {}): Promise<ValidationResult> {
  const officeVersion = options.officeVersion || 'Office2019';
  const backend = getValidatorBackend();
  return backend.validateFile(filePath, { officeVersion });
}

/**
 * Asserts that a DOCX buffer is schema-valid. Throws OoxmlValidationError if invalid.
 */
export async function assertValidDocx(buffer: Buffer, context: string = 'DOCX export', options: ValidateOptions = {}): Promise<void> {
  const res = await validateDocxBuffer(buffer, options);
  if (!res.ok) {
    throw new OoxmlValidationError(res.errors, context);
  }
}
