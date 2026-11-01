/**
 * Text of a stored XML part. Parts the package declares in UTF-16 (a byte
 * order mark, or the first characters of an XML declaration as 16-bit units)
 * are decoded as UTF-16; everything else is read as UTF-8 exactly as before.
 * A UTF-16 part with an odd number of bytes is not decodable and is read as
 * UTF-8, which the well-formedness check then refuses.
 */
export function decodeXmlPart(data: Buffer): string {
  const n = data.length;
  if (n >= 2 && n % 2 === 0) {
    if (data[0] === 0xff && data[1] === 0xfe) return data.subarray(2).toString('utf16le');
    if (data[0] === 0xfe && data[1] === 0xff) return swapBytes(data.subarray(2)).toString('utf16le');
    if (n >= 4 && data[0] === 0x3c && data[1] === 0x00 && data[2] === 0x3f && data[3] === 0x00) return data.toString('utf16le');
    if (n >= 4 && data[0] === 0x00 && data[1] === 0x3c && data[2] === 0x00 && data[3] === 0x3f) return swapBytes(data).toString('utf16le');
  }
  return data.toString('utf8');
}

function swapBytes(data: Buffer): Buffer {
  const out = Buffer.from(data);
  out.swap16();
  return out;
}
