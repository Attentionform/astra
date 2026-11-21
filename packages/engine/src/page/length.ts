import { toError } from '../errors';

/**
 * A length as a caller writes it: a number, which is twips (a twentieth of a point, 1440 to an inch), or a string with a
 * unit such as `"25.4mm"`, `"1in"`, `"2.5cm"`, `"72pt"` or `"1440tw"`.
 */
export type LengthInput = number | string;

const TWIPS_PER_UNIT: Record<string, number> = {
  tw: 1,
  twip: 1,
  twips: 1,
  pt: 20,
  in: 1440,
  inch: 1440,
  inches: 1440,
  '"': 1440,
  mm: 1440 / 25.4,
  cm: 1440 / 2.54,
};

/** The units a length may use, for messages and docs. */
export const LENGTH_UNITS = ['mm', 'cm', 'in', 'pt', 'tw'] as const;

const LENGTH_PATTERN = /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*([a-zA-Z"]*)\s*$/;

/** Splits `"25.4mm"` into its number and unit. The unit is empty for a bare number. Null when the text is not a length. */
export function splitLength(text: string): { value: number; unit: string } | null {
  const m = LENGTH_PATTERN.exec(text);
  if (!m) return null;
  return { value: Number(m[1]), unit: m[2].toLowerCase() };
}

/**
 * Converts a length to whole twips, rounding to the nearest twip. `what` names the field in the error message.
 * A number or a string without a unit is twips. Throws `E_VALIDATION` for anything that is not a finite length.
 */
export function parseLength(input: unknown, what: string): number {
  if (typeof input === 'number') {
    if (!Number.isFinite(input)) {
      throw toError('E_VALIDATION', `${what} must be a finite number, got ${String(input)}`, `pass a number of twips or a length such as "25mm" or "1in"`);
    }
    return Math.round(input);
  }
  if (typeof input !== 'string') {
    throw toError('E_VALIDATION', `${what} must be a number or a length string, got ${input === null ? 'null' : typeof input}`, `pass a number of twips or a length such as "25mm" or "1in"`);
  }
  const parts = splitLength(input);
  if (!parts) {
    throw toError('E_VALIDATION', `${what} '${input}' is not a length`, `write a number followed by a unit (${LENGTH_UNITS.join(', ')}), for example "25mm" or "1in"; a number alone is twips`);
  }
  const perUnit = parts.unit === '' ? 1 : TWIPS_PER_UNIT[parts.unit];
  if (perUnit === undefined) {
    throw toError('E_VALIDATION', `${what} '${input}' has an unknown unit '${parts.unit}'`, `use one of ${LENGTH_UNITS.join(', ')}`);
  }
  return Math.round(parts.value * perUnit);
}

/** Twips to millimetres, to two decimals. For reports. */
export const twipsToMm = (twips: number): number => Math.round((twips / 1440) * 25.4 * 100) / 100;

/** Twips to inches, to three decimals. For reports. */
export const twipsToInches = (twips: number): number => Math.round((twips / 1440) * 1000) / 1000;
