/** Characters XML 1.0 cannot hold: C0 controls other than tab, line feed and carriage return, and U+FFFE / U+FFFF. */
// eslint-disable-next-line no-control-regex -- intentional: the whole point is to match the C0 controls XML 1.0 forbids
const XML_ILLEGAL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/** Drop the characters XML 1.0 cannot hold, so a name from an option or the environment cannot make a part unreadable. */
export function stripIllegalXmlChars(value: string): string {
  return value.replace(XML_ILLEGAL_CHARS, '');
}
