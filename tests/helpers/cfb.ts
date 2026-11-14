/** Builds a small compound file (OLE container) with the given directory entry names, for the container tests. */
export function buildCfb(names: string[], opts: { loopDirectory?: boolean; truncateTo?: number; sectorShift?: 9 | 12 } = {}): Buffer {
  const shift = opts.sectorShift ?? 9;
  const sectorSize = 1 << shift;
  const header = Buffer.alloc(sectorSize, 0);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(header, 0);
  header.writeUInt16LE(0x3e, 0x18); // minor version
  header.writeUInt16LE(shift === 9 ? 3 : 4, 0x1a); // major version
  header.writeUInt16LE(0xfffe, 0x1c); // byte order
  header.writeUInt16LE(shift, 0x1e);
  header.writeUInt32LE(1, 0x2c); // one FAT sector
  header.writeUInt32LE(1, 0x30); // directory starts at sector 1
  header.writeUInt32LE(0xfffffffe, 0x44); // no DIFAT chain
  header.writeUInt32LE(0, 0x4c); // FAT is sector 0
  for (let i = 1; i < 109; i++) header.writeUInt32LE(0xffffffff, 0x4c + i * 4);

  const perSector = sectorSize / 128;
  const all = ['Root Entry', ...names];
  const count = Math.max(1, Math.ceil(all.length / perSector));
  if (count + 1 > sectorSize / 4) throw new Error('too many entries for this helper');
  const fat = Buffer.alloc(sectorSize, 0xff);
  fat.writeUInt32LE(0xfffffffd, 0); // sector 0 is the FAT itself
  for (let i = 0; i < count; i++) {
    const last = i === count - 1;
    fat.writeUInt32LE(last ? (opts.loopDirectory ? 1 : 0xfffffffe) : i + 2, (i + 1) * 4); // directory sectors 1..count, chained
  }
  const dirSectors: Buffer[] = [];
  for (let n = 0; n < count; n++) dirSectors.push(Buffer.alloc(sectorSize, 0));
  all.forEach((name, i) => {
    const dir = dirSectors[Math.floor(i / perSector)];
    const off = (i % perSector) * 128;
    const raw = Buffer.from(`${name}\0`, 'utf16le');
    raw.copy(dir, off);
    dir.writeUInt16LE(raw.length, off + 64);
    dir[off + 66] = i === 0 ? 5 : 2; // root storage, then streams
  });
  const file = Buffer.concat([header, fat, ...dirSectors]);
  return opts.truncateTo !== undefined ? file.subarray(0, opts.truncateTo) : file;
}
