/**
 * OOXML Validation Gate using @attentionform/ooxml-validator (.NET OpenXmlValidator)
 * 
 * Validates .docx buffers or files against official ECMA-376 / ISO-29500 XSD schemas.
 * Rejects corrupt documents before they are saved to disk or delivered to users.
 */

import { getValidatorBackend, type ValidationResult, type ValidateOptions } from './validator-backend';
import { resolveOfficeVersion } from './office-version';

export type { ValidationResult, ValidateOptions };
export * from './office-version';

/**
 * One error in a validation result: a `description`, the `path` of the part, the `xPath` of the element, the SDK `id` and the `errorType`.
 */
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

/**
 * Thrown by `assertValidDocx`, and by PDF export, when a document fails Office schema validation.
 * `errors` lists every schema error and `firstError` is the first one. `code` is `E_VALIDATION`. `remedy` says what to do next.
 */
export class OoxmlValidationError extends Error {
  public readonly errors: ValidationError[];
  public readonly firstError: ValidationError;
  public readonly code = 'E_VALIDATION' as const;
  public readonly remedy: string;
  public readonly retryable = false as const;

  /** `remedy` replaces the default text for callers that know a better next step (PDF export does). */
  constructor(errors: ValidationError[], context: string = 'Document', remedy?: string) {
    const first = errors[0] || {};
    const xp = (typeof first.xPath === 'string' && first.xPath) || (typeof first.path === 'string' && first.path) || 'unknown';
    const hint = tableIndexHintForXPath((first as { xPath?: unknown }).xPath ?? (first as { path?: unknown }).path);
    const detail = `${first.description || 'Schema error'} @ ${xp}${hint ? ` (${hint})` : ''}`;
    super(`[OOXML Schema Violation] ${context} has ${errors.length} validation error(s): ${detail}`);
    this.remedy = remedy ?? 're-export the file from Word (or pick a valid .docx) and retry';
    this.name = 'OoxmlValidationError';
    this.errors = errors;
    this.firstError = first;
  }
}

/**
 * Checks the bytes of a .docx file against the Office schema. The target version is `options.officeVersion`, else `ASTRA_OFFICE_VERSION`, else Office 2019.
 * It returns the errors that it found. A document that fails the schema is a result, not a thrown error.
 */
export async function validateDocxBuffer(buffer: Buffer, options: ValidateOptions = {}): Promise<ValidationResult> {
  const backend = getValidatorBackend();
  // One explicit version goes to the validator: the explicit option, ASTRA_OFFICE_VERSION, then Office2019.
  const officeVersion = resolveOfficeVersion(options.officeVersion);
  const result = await backend.validateBuffer(buffer, { ...options, officeVersion });
  return { ...result, requestedOfficeVersion: officeVersion, validationStatus: 'completed' };
}

/**
 * Checks a .docx file on disk against the Office schema, as `validateDocxBuffer` does.
 */
export async function validateDocxFile(filePath: string, options: ValidateOptions = {}): Promise<ValidationResult> {
  const officeVersion = resolveOfficeVersion(options.officeVersion);
  const backend = getValidatorBackend();
  const result = await backend.validateFile(filePath, { officeVersion });
  return { ...result, requestedOfficeVersion: officeVersion, validationStatus: 'completed' };
}

/**
 * Throws `OoxmlValidationError` when the bytes of a .docx file fail the Office schema, and returns nothing when they pass. `context` names the file or step in the message.
 */
export async function assertValidDocx(buffer: Buffer, context: string = 'DOCX export', options: ValidateOptions = {}): Promise<void> {
  const res = await validateDocxBuffer(buffer, options);
  if (!res.ok) {
    throw new OoxmlValidationError(res.errors, context);
  }
}
