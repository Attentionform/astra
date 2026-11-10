/**
 * 'start' and 'end' (also spelled '@start' and '@end') stand for the first and last block of the
 * document. Exact lowercase match only, so a real node id such as 'End' passes through untouched.
 */
export function anchorAliasOf(value: unknown): 'start' | 'end' | null {
  if (value === 'start' || value === '@start') return 'start';
  if (value === 'end' || value === '@end') return 'end';
  return null;
}
