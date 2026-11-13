/**
 * Counts dark pixels in a PNG (8-bit gray, gray+alpha, RGB or RGBA, not interlaced), with no
 * dependencies. A render of a page with text has ink; a blank page has none.
 */
import { inflateSync } from 'node:zlib';

const SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

export function isPng(bytes) {
  return bytes.length >= 8 && bytes.subarray(0, 8).equals(SIGNATURE);
}

export function pngInfo(bytes) {
  if (!isPng(bytes)) throw new Error('not a PNG');
  return {
    width: bytes.readUInt32BE(16),
    height: bytes.readUInt32BE(20),
    bitDepth: bytes[24],
    colorType: bytes[25],
    interlace: bytes[28],
  };
}

export function countInk(bytes, darkBelow = 128) {
  const info = pngInfo(bytes);
  const channels = { 0: 1, 2: 3, 4: 2, 6: 4 }[info.colorType];
  if (!channels || info.bitDepth !== 8 || info.interlace !== 0) {
    throw new Error(`unsupported PNG (colorType ${info.colorType}, bitDepth ${info.bitDepth}, interlace ${info.interlace})`);
  }
  const idat = [];
  for (let at = 8; at + 8 <= bytes.length; ) {
    const length = bytes.readUInt32BE(at);
    const type = bytes.toString('latin1', at + 4, at + 8);
    if (type === 'IDAT') idat.push(bytes.subarray(at + 8, at + 8 + length));
    at += 12 + length;
  }
  const raw = inflateSync(Buffer.concat(idat));
  const stride = info.width * channels;
  const rows = [];
  let previous = Buffer.alloc(stride);
  for (let y = 0; y < info.height; y++) {
    const filter = raw[y * (stride + 1)];
    const line = Buffer.from(raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1)));
    for (let x = 0; x < stride; x++) {
      const left = x >= channels ? line[x - channels] : 0;
      const up = previous[x];
      const upLeft = x >= channels ? previous[x - channels] : 0;
      let add = 0;
      if (filter === 1) add = left;
      else if (filter === 2) add = up;
      else if (filter === 3) add = (left + up) >> 1;
      else if (filter === 4) {
        const p = left + up - upLeft;
        const pa = Math.abs(p - left);
        const pb = Math.abs(p - up);
        const pc = Math.abs(p - upLeft);
        add = pa <= pb && pa <= pc ? left : pb <= pc ? up : upLeft;
      } else if (filter !== 0) throw new Error(`bad PNG filter ${filter}`);
      line[x] = (line[x] + add) & 0xff;
    }
    rows.push(line);
    previous = line;
  }
  let ink = 0;
  for (const line of rows) {
    for (let x = 0; x < info.width; x++) {
      const base = x * channels;
      const luminance =
        info.colorType === 0 || info.colorType === 4
          ? line[base]
          : (line[base] * 299 + line[base + 1] * 587 + line[base + 2] * 114) / 1000;
      if (luminance < darkBelow) ink++;
    }
  }
  return { ...info, ink };
}
