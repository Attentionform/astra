/**
 * The CLI side of the baseline-aware validation gate. A file that already has validator errors can be
 * written when an edit adds none; `--strict` refuses every error as before.
 */
import {
  BASELINE_MAX_ERRORS,
  buildBaseline,
  judgeAgainstBaseline,
  validateDocxBuffer,
  type BaselineVerdict,
  type ValidationResult,
} from '@attentionform/astra-engine';

export { BASELINE_MAX_ERRORS };

export type ToleratedVerdict = Extract<BaselineVerdict, { kind: 'tolerated' }>;
export type IntroducedVerdict = Extract<BaselineVerdict, { kind: 'introduced' }>;

/**
 * Judges the result of a gate. A valid result is clean. Otherwise, unless `strict`, the errors are
 * compared with those of the original bytes, which are validated only now and only when needed.
 */
export async function judgeGate(result: ValidationResult, strict: boolean, original: () => Buffer): Promise<BaselineVerdict> {
  if (result.ok) return { kind: 'clean' };
  if (strict) return { kind: 'incomparable', reason: 'no-baseline' };
  let baseline = null;
  try {
    baseline = buildBaseline(await validateDocxBuffer(original(), { maxErrors: BASELINE_MAX_ERRORS }));
  } catch {
    baseline = null;
  }
  return judgeAgainstBaseline(result, baseline);
}

/** True when the gate must refuse the write. */
export function refuses(verdict: BaselineVerdict): boolean {
  return verdict.kind !== 'clean' && verdict.kind !== 'tolerated';
}

/** Wording of a refusal for an edit that adds errors to a file that already had some. */
export function introducedText(verdict: IntroducedVerdict, description: string): string {
  return `this edit adds ${verdict.introduced.length} validator error(s) (the file already had ${verdict.preExisting}), nothing written: ${description}.`;
}

/** Fields added to the JSON output when errors the file already had were tolerated. */
export function toleratedFields(verdict: ToleratedVerdict | undefined): Record<string, unknown> {
  return verdict ? { validation: { tolerated: true, preExistingErrors: verdict.preExisting, introducedErrors: 0 } } : {};
}

/** The line written to stderr when errors the file already had were tolerated. */
export function toleratedNote(verdict: ToleratedVerdict): string {
  return `saved with ${verdict.preExisting} existing validator errors, 0 introduced. The validator is stricter than Word; Word may still open the file.`;
}
