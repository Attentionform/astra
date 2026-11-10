/**
 * Shared pieces of the baseline-aware validation gate: the server settings, the baseline of a
 * session, and the wording of the answers. A document that already has validator errors can be
 * saved when an edit adds none; the answer says how many errors were there before.
 */
import * as crypto from 'crypto';
import * as fsp from 'fs/promises';
import {
  BASELINE_MAX_ERRORS,
  buildBaseline,
  judgeAgainstBaseline,
  validateDocxBuffer,
  type BaselineVerdict,
  type ValidationBaseline,
  type ValidationResult,
} from '@attentionform/astra-engine';
import type { DocHandle } from '../session-store';
import { COMMIT_MESSAGE_MAX } from './errors';
import { remedyForValidationError } from './diff';

const DEFAULT_OPEN_VALIDATION_TIMEOUT_MS = 5000;

let strictOverride: boolean | undefined;
let openTimeoutOverride: number | undefined;

/** Set by the `--strict-validation` flag. */
export function setStrictValidationDefault(value: boolean): void {
  strictOverride = value;
}

/** True when a save refuses every validator error, whether or not the file had it before. */
export function strictValidationDefault(): boolean {
  if (strictOverride !== undefined) return strictOverride;
  return process.env.ASTRA_STRICT_VALIDATION === '1';
}

/** Set by the `--open-validation-timeout` flag. */
export function setOpenValidationTimeoutMs(value: number): void {
  openTimeoutOverride = value;
}

/** How long open_document waits for the validator, in milliseconds. 0 skips the check. */
export function openValidationTimeoutMs(): number {
  if (openTimeoutOverride !== undefined) return openTimeoutOverride;
  const raw = process.env.ASTRA_OPEN_VALIDATION_TIMEOUT_MS;
  if (raw !== undefined && /^\d+$/.test(raw)) return Number(raw);
  return DEFAULT_OPEN_VALIDATION_TIMEOUT_MS;
}

export function sha256Hex(bytes: Buffer): string {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

/** Validates a package with the full error list, so the result can be compared. */
export function validateForBaseline(bytes: Buffer): Promise<ValidationResult> {
  return validateDocxBuffer(bytes, { maxErrors: BASELINE_MAX_ERRORS });
}

/** Stores the baseline of a package the session now holds as its saved state. */
export function adoptBaseline(handle: DocHandle, result: ValidationResult): void {
  handle.validationBaseline = buildBaseline(result);
}

/**
 * The baseline of a session. When the open did not produce one (timed out or skipped), the file on
 * disk is validated, but only when it is still the file that was opened. Otherwise null, and the
 * save is judged as not comparable.
 */
export async function resolveBaseline(handle: DocHandle): Promise<ValidationBaseline | null> {
  if (handle.validationBaseline !== undefined) return handle.validationBaseline;
  if (!handle.openSha256) return null;
  try {
    const bytes = await fsp.readFile(handle.filePath);
    if (sha256Hex(bytes) !== handle.openSha256) return null;
    const baseline = buildBaseline(await validateForBaseline(bytes));
    handle.validationBaseline = baseline;
    return baseline;
  } catch {
    return null;
  }
}

/**
 * Verdict for a package that revert_document or redo_document restores. Such a state was saved before,
 * so it is allowed to carry the errors the file had when it was opened, as well as the errors of the
 * saved state of the session. Anything beyond both is refused.
 */
export async function judgeRestoredPackage(handle: DocHandle, result: ValidationResult): Promise<BaselineVerdict> {
  if (result.ok) return { kind: 'clean' };
  if (strictValidationDefault()) return { kind: 'incomparable', reason: 'no-baseline' };
  const current = await resolveBaseline(handle);
  const opened = handle.openBaseline ?? null;
  if (current === null && opened === null) return { kind: 'incomparable', reason: 'no-baseline' };
  const counts: Record<string, number> = { ...(current?.counts ?? {}) };
  for (const [k, v] of Object.entries(opened?.counts ?? {})) counts[k] = Math.max(counts[k] ?? 0, v);
  const total = Object.values(counts).reduce((a, b) => a + b, 0);
  return judgeAgainstBaseline(result, { counts, total });
}

interface MessageSource {
  description?: unknown;
  message?: unknown;
}

function descriptionOf(e: unknown): string {
  const m = e as MessageSource | undefined;
  return (typeof m?.description === 'string' && m.description) || (typeof m?.message === 'string' && m.message) || 'OOXML validation failed';
}

/** The first `limit` descriptions of a result, cut to a readable length. */
export function firstMessages(result: ValidationResult, limit = 5): string[] {
  return (result.errors ?? []).slice(0, limit).map((e) => descriptionOf(e).slice(0, 300));
}

/** Validation errors carried on an answer (bounded, first errors only). */
export const AUDIT_VALIDATION_ERRORS_MAX = 10;

export interface AuditValidation {
  valid: boolean;
  errorCount: number;
  errors: Array<{ description: string; xPath?: string }>;
}

/** The first `max` errors as {description, xPath?}, each cut to a readable length. */
export function summarizeValidationErrors(errors: unknown, max: number = AUDIT_VALIDATION_ERRORS_MAX): Array<{ description: string; xPath?: string }> {
  const raw = Array.isArray(errors) ? errors.slice(0, max) : [];
  return raw.map((e) => {
    const err = e as { description?: unknown; message?: unknown; xPath?: unknown; path?: unknown };
    const description = String(
      (typeof err.description === 'string' && err.description) ||
        (typeof err.message === 'string' && err.message) ||
        'OOXML validation error'
    ).slice(0, 500);
    const xp = typeof err.xPath === 'string' ? err.xPath : typeof err.path === 'string' ? err.path : undefined;
    return xp ? { description, xPath: xp } : { description };
  });
}

/** The shape validate_document returns: valid, how many errors, and the first of them. */
export function toAuditValidation(
  validation: { ok: boolean; errors?: unknown },
  max: number = AUDIT_VALIDATION_ERRORS_MAX
): AuditValidation {
  return {
    valid: validation.ok,
    errorCount: Array.isArray(validation.errors) ? validation.errors.length : 0,
    errors: summarizeValidationErrors(validation.errors, max),
  };
}

/** The error a save or normalize raises when the edit adds validator errors. */
export function introducedError(
  toolName: string,
  verdict: Extract<BaselineVerdict, { kind: 'introduced' }>,
): Error {
  const first = verdict.introduced[0];
  const desc = descriptionOf(first);
  const xp = typeof first?.xPath === 'string' ? first.xPath : undefined;
  return Object.assign(
    new Error(
      `E_VALIDATION: ${toolName} blocked: this edit adds ${verdict.introduced.length} validator error(s) (the file already had ${verdict.preExisting}): ${desc.slice(0, 2000)}`,
    ),
    {
      code: 'E_VALIDATION',
      remedy: remedyForValidationError(xp, desc),
      errorCount: verdict.introduced.length,
      errors: summarizeValidationErrors(verdict.introduced),
    },
  );
}

/** The `validation` field of an answer for a result that was judged. */
export function validationField(result: ValidationResult, verdict: BaselineVerdict) {
  if (verdict.kind === 'tolerated') {
    return {
      ok: false as const,
      tolerated: true as const,
      preExistingErrors: verdict.preExisting,
      introducedErrors: 0 as const,
      messages: firstMessages(result),
      valid: false as const,
      errorCount: Array.isArray(result.errors) ? result.errors.length : verdict.preExisting,
    };
  }
  // `valid` and `errorCount` sit next to `ok` so a save reads like validate_document's answer.
  return { ok: true as const, valid: true as const, errorCount: 0 as const };
}

/** Adds the note about existing errors to a commit message; the message is cut, the note never. */
export function withToleratedNote(message: string, preExisting: number): string {
  const note = ` (saved with ${preExisting} existing validator errors, 0 introduced)`;
  const room = Math.max(0, COMMIT_MESSAGE_MAX - note.length);
  return (message.length > room ? message.slice(0, room) : message) + note;
}

export type OpenValidationField =
  | { status: 'valid' }
  | {
      status: 'invalid';
      preExistingErrors: number;
      baselineComparable: boolean;
      messages: string[];
      degraded: true;
      note: string;
    }
  | { status: 'timed-out'; degraded: 'unknown'; note: string }
  | { status: 'skipped'; note?: string };

/**
 * Validates the file that was just opened and stores its baseline on the handle. The open itself is
 * never refused for these reasons, and nothing is reported as valid unless the validator said so.
 */
export async function validateAtOpen(handle: DocHandle, timeoutMs: number): Promise<OpenValidationField> {
  let bytes: Buffer;
  try {
    bytes = await fsp.readFile(handle.filePath);
  } catch (err: unknown) {
    return { status: 'skipped', note: `the file could not be read for validation: ${err instanceof Error ? err.message : String(err)}` };
  }
  // Kept even when the check is skipped, so the first save can validate the opened file if it is still unchanged.
  handle.openSha256 = sha256Hex(bytes);
  if (timeoutMs === 0) return { status: 'skipped' };
  const pending = validateForBaseline(bytes);
  // A late result of a check that was given up on is dropped, but must not become an unhandled rejection.
  pending.catch(() => undefined);
  let timer: NodeJS.Timeout | undefined;
  const timeout = new Promise<'timeout'>((resolve) => {
    timer = setTimeout(() => resolve('timeout'), timeoutMs);
  });
  let result: ValidationResult | 'timeout';
  try {
    result = await Promise.race([pending, timeout]);
  } catch (err: unknown) {
    return { status: 'skipped', note: `the validator is not available: ${err instanceof Error ? err.message : String(err)}` };
  } finally {
    if (timer) clearTimeout(timer);
  }
  if (result === 'timeout') {
    return {
      status: 'timed-out',
      degraded: 'unknown',
      note: `validation did not finish within ${timeoutMs} ms; the file may have errors; they will be checked at first save`,
    };
  }
  const baseline = buildBaseline(result);
  handle.validationBaseline = baseline;
  handle.openBaseline = baseline;
  if (result.ok) return { status: 'valid' };
  const comparable = baseline !== null;
  const stricter = 'The validator is stricter than Word, so Word may still open the file.';
  return {
    status: 'invalid',
    preExistingErrors: result.errors?.length ?? result.errorCount ?? 0,
    baselineComparable: comparable,
    messages: firstMessages(result),
    degraded: true,
    note: comparable
      ? `${stricter} Edits that add no new error can be saved.`
      : `${stricter} The error list could not be compared (it was cut short or holds an exception), so every error is refused at save.`,
  };
}
