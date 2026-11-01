/**
 * Baseline-aware validation. A document can arrive with validator errors that it already had before
 * any edit (the validator is stricter than Word). The baseline records those errors by signature, so a
 * later save can tell an error the edit added from an error that was there already.
 */
import type { ValidationResult } from './validator-backend';

/** Every baseline-aware validation asks for this many errors, so the list is not cut at the default of 200. */
export const BASELINE_MAX_ERRORS = 5000;

interface BaselineError {
  description?: string | null;
  path?: string | null;
  xPath?: string | null;
  id?: string | null;
  errorType?: string | null;
  part?: string | null;
}

export interface ValidationBaseline {
  /** Signature of an error to the number of errors with that signature. */
  counts: Record<string, number>;
  total: number;
}

export type BaselineVerdict =
  | { kind: 'clean' }
  | { kind: 'tolerated'; preExisting: number }
  | { kind: 'introduced'; introduced: BaselineError[]; preExisting: number }
  | { kind: 'incomparable'; reason: 'truncated' | 'exception' | 'guard' | 'no-baseline' };

/** Part, validator id and xPath with every index removed, so moving an element does not change it. */
export function errorSignature(e: BaselineError): string {
  return `${e.part ?? e.path ?? ''}|${e.id ?? ''}|${(e.xPath ?? '').replace(/\[\d+\]/g, '')}`;
}

/** Why a result cannot be compared, or null when it can. */
function incomparableReason(result: ValidationResult): 'truncated' | 'exception' | 'guard' | null {
  if (result.truncated) return 'truncated';
  const errors = (result.errors ?? []) as BaselineError[];
  if (errors.some((e) => e.errorType === 'Exception')) return 'exception';
  if (errors.some((e) => e.errorType === 'DuplicateParaId')) return 'guard';
  return null;
}

/** The baseline of a result, or null when the result is cut short or holds an exception or guard error. */
export function buildBaseline(result: ValidationResult): ValidationBaseline | null {
  if (incomparableReason(result)) return null;
  const counts: Record<string, number> = {};
  const errors = (result.errors ?? []) as BaselineError[];
  for (const e of errors) {
    const sig = errorSignature(e);
    counts[sig] = (counts[sig] ?? 0) + 1;
  }
  return { counts, total: errors.length };
}

/** Decides whether the errors in `result` were all there before (`tolerated`) or some are new (`introduced`). */
export function judgeAgainstBaseline(result: ValidationResult, baseline: ValidationBaseline | null): BaselineVerdict {
  if (result.ok) return { kind: 'clean' };
  const reason = incomparableReason(result);
  if (reason) return { kind: 'incomparable', reason };
  if (!baseline) return { kind: 'incomparable', reason: 'no-baseline' };
  const seen: Record<string, number> = {};
  const introduced: BaselineError[] = [];
  const errors = (result.errors ?? []) as BaselineError[];
  for (const e of errors) {
    const sig = errorSignature(e);
    seen[sig] = (seen[sig] ?? 0) + 1;
    if (seen[sig] > (baseline.counts[sig] ?? 0)) introduced.push(e);
  }
  if (introduced.length === 0) return { kind: 'tolerated', preExisting: errors.length };
  return { kind: 'introduced', introduced, preExisting: baseline.total };
}
