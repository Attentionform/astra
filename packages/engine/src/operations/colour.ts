/**
 * Colour values in operations: one rule for every field that takes a colour.
 *
 * Accepted forms: six hex digits with an optional leading `#`, optionally the word `auto` (only where the
 * caller allows it), and ten common names mapped to a fixed hex value. Surrounding whitespace is ignored.
 * The writers call `normalizeColour`, and the schema calls it too, so the set of values the schema accepts
 * and the set the writers accept cannot drift apart. The schema is only a gate: the operations a writer
 * receives are the caller's own, not the schema's output, so a trim or a mapping done only in the schema
 * would never reach the writer.
 *
 * Name values are the CSS basic colour keywords, not Word's palette (Word's own "green" is 00B050).
 */

export const COLOUR_NAMES: Readonly<Record<string, string>> = {
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  black: '000000',
  white: 'FFFFFF',
  yellow: 'FFFF00',
  orange: 'FFA500',
  gray: '808080',
  grey: '808080',
  purple: '800080',
};

const HEX = /^#?([0-9A-Fa-f]{6})$/;

export interface ColourOptions {
  /** Accept the OOXML value `auto` (the application picks the colour). */
  allowAuto?: boolean;
}

/**
 * The value to write for `input`, or undefined when `input` is not an accepted colour. Hex keeps its case
 * (the writers have always written it as given), a name becomes its uppercase hex, `auto` stays `auto`.
 * Names and `auto` are matched case-insensitively.
 */
export function normalizeColour(input: unknown, options: ColourOptions = {}): string | undefined {
  if (typeof input !== 'string') return undefined;
  const text = input.trim();
  const hex = HEX.exec(text);
  if (hex) return hex[1];
  const lower = text.toLowerCase();
  if (options.allowAuto && lower === 'auto') return 'auto';
  return Object.prototype.hasOwnProperty.call(COLOUR_NAMES, lower) ? COLOUR_NAMES[lower] : undefined;
}

/**
 * For writers that keep their own strict check after this: the mapped or trimmed value when `input` is an accepted
 * colour, else `input` unchanged (so the writer's own check still reports it).
 */
export function resolveColourInput<T>(input: T, options: ColourOptions = {}): T | string {
  const value = normalizeColour(input, options);
  return value === undefined ? input : value;
}

/** The accepted forms, for an error message. */
export function colourForms(options: ColourOptions = {}): string {
  const names = Object.keys(COLOUR_NAMES).join(', ');
  return `6 hex digits such as FF0000 (optionally with #), ${options.allowAuto ? '"auto", ' : ''}or one of: ${names}`;
}

/** E_VALIDATION error for a colour that no form accepts. */
export function colourError(field: string, value: unknown, options: ColourOptions = {}): Error {
  return new Error(
    `E_VALIDATION: ${field} must be ${colourForms(options)}. Got ${JSON.stringify(value)}. ` +
      `Remedy: pass ${field} as hex, for example "FF0000", or a colour name such as "red".`
  );
}

/** `normalizeColour`, or an E_VALIDATION error naming the field. For writers. */
export function requireColour(field: string, input: unknown, options: ColourOptions = {}): string {
  const value = normalizeColour(input, options);
  if (value === undefined) throw colourError(field, input, options);
  return value;
}

/** A border width is a whole number of eighths of a point, from 1 (an eighth) to 96 (12 pt). */
export const BORDER_WIDTH_MIN = 1;
export const BORDER_WIDTH_MAX = 96;

/**
 * The border width in `input` (a number, or a string holding a whole number), or `fallback` when it is undefined.
 * Anything else, or a width outside BORDER_WIDTH_MIN..BORDER_WIDTH_MAX, is an E_VALIDATION error naming the field.
 * The schema checks the top-level `sz`; this is the same rule for the writers, which also read it from the nested
 * `params` of a table operation and from the border configs of a style operation.
 */
export function requireBorderSize(field: string, input: unknown, fallback: number): number {
  if (input === undefined) return fallback;
  const n = typeof input === 'number' ? input : typeof input === 'string' && input.trim() !== '' ? Number(input) : Number.NaN;
  if (!Number.isInteger(n) || n < BORDER_WIDTH_MIN || n > BORDER_WIDTH_MAX) {
    throw new Error(
      `E_VALIDATION: ${field}: border width must be a whole number of eighths of a point from ${BORDER_WIDTH_MIN} to ${BORDER_WIDTH_MAX} (got ${JSON.stringify(input)}). ` +
        'Remedy: pass sz:4 for a 0.5 pt line, sz:8 for 1 pt, sz:96 for 12 pt (the widest).'
    );
  }
  return n;
}

/** JSON-schema `pattern` that accepts exactly what `normalizeColour` accepts (case-insensitive forms expanded). */
export function colourPattern(options: ColourOptions = {}): string {
  const ci = (word: string): string => [...word].map((c) => `[${c.toLowerCase()}${c.toUpperCase()}]`).join('');
  const words = [...Object.keys(COLOUR_NAMES), ...(options.allowAuto ? ['auto'] : [])].map(ci);
  return `^\\s*(#?[0-9A-Fa-f]{6}|${words.join('|')})\\s*$`;
}

/** The `description` of a colour field in the generated schemas. */
export function colourDescription(options: ColourOptions = {}): string {
  return `A colour: ${colourForms(options)}.`;
}

/** The highlight values of ECMA-376 ST_HighlightColor, plus `none` (remove the highlight). */
export const HIGHLIGHT_COLOURS = [
  'yellow', 'green', 'cyan', 'magenta', 'blue', 'red',
  'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow',
  'darkGray', 'lightGray', 'black', 'white', 'none',
] as const;
