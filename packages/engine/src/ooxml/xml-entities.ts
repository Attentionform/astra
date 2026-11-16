/**
 * Decoding for text read out of XML with a regular expression. A parser decodes attribute values before the
 * caller sees them; a regex does not, so `Target="../embeddings/w&#49;.xlsx"` and `Target="../embeddings/w1.xlsx"`
 * look different even though they name the same part.
 */

const NAMED_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

/** Decodes the five predefined entities and decimal / hexadecimal character references in one pass. */
export function decodeXmlEntities(value: string): string {
  return value.replace(/&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|([A-Za-z]+));/g, (whole, dec?: string, hex?: string, name?: string) => {
    if (name !== undefined) return NAMED_ENTITIES[name] ?? whole;
    const code = dec !== undefined ? Number.parseInt(dec, 10) : Number.parseInt(hex as string, 16);
    if (!Number.isFinite(code) || code < 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) return whole;
    return String.fromCodePoint(code);
  });
}

/**
 * A relationship target is a URI reference, so a percent-escape names the same character as the plain form
 * (`w%31.xlsx` is `w1.xlsx`). Returns the input when it is not valid percent-encoding.
 */
export function decodePartUri(target: string): string {
  try {
    return decodeURIComponent(target);
  } catch {
    return target;
  }
}
