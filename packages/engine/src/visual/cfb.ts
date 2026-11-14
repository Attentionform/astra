/**
 * Just enough of the Compound File Binary format (the OLE container) to tell a password-protected Office Open XML
 * file from every other thing that is not a zip package.
 *
 * Word wraps an encrypted .docx in a compound file that holds only `EncryptionInfo`, `EncryptedPackage` and a few
 * `DataSpaces` streams. Its content cannot be read without the password, so no renderer can follow a reference in
 * it. A legacy .doc (a `WordDocument` stream), a spreadsheet or anything else in a compound file CAN carry
 * references, so those stay unverifiable. The check therefore reads the directory properly and accepts only a
 * file whose every entry name is one of the known ones. Anything unexpected, truncated or odd returns false.
 */

const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const END_OF_CHAIN = 0xfffffffe;
const FREE_SECTOR = 0xffffffff;
const MAX_SECTORS = 1 << 20;

/** Entry names of an encrypted OOXML package (ECMA-376 part 1, 2.3.4 and MS-OFFCRYPTO). Compared without case. */
const ENCRYPTED_PACKAGE_NAMES = new Set(
  [
    'root entry',
    'encryptioninfo',
    'encryptedpackage',
    '\u0006dataspaces',
    'version',
    'dataspacemap',
    'dataspaceinfo',
    'strongencryptiondataspace',
    'transforminfo',
    'strongencryptiontransform',
    '\u0006primary',
  ]
);

/** The names of every directory entry, or null when the container is not a well-formed compound file. */
export function listCfbEntryNames(buf: Buffer): string[] | null {
  try {
    if (buf.length < 512 || !buf.subarray(0, 8).equals(CFB_MAGIC)) return null;
    const major = buf.readUInt16LE(0x1a);
    const shift = buf.readUInt16LE(0x1e);
    if ((major !== 3 && major !== 4) || (shift !== 9 && shift !== 12)) return null;
    const sectorSize = 1 << shift;
    const sectorAt = (n: number): Buffer | null => {
      const start = (n + 1) * sectorSize;
      return n >= 0 && n < MAX_SECTORS && start + sectorSize <= buf.length ? buf.subarray(start, start + sectorSize) : null;
    };
    const perSector = sectorSize / 4;
    const fatSectorCount = buf.readUInt32LE(0x2c);
    const firstDirSector = buf.readUInt32LE(0x30);
    if (fatSectorCount === 0 || fatSectorCount > MAX_SECTORS) return null;

    // The sectors that hold the allocation table: 109 in the header, the rest through the DIFAT chain.
    const fatSectors: number[] = [];
    for (let i = 0; i < 109 && fatSectors.length < fatSectorCount; i++) {
      const s = buf.readUInt32LE(0x4c + i * 4);
      if (s !== FREE_SECTOR) fatSectors.push(s);
    }
    let difat = buf.readUInt32LE(0x44);
    for (let hops = 0; fatSectors.length < fatSectorCount && difat !== END_OF_CHAIN && difat !== FREE_SECTOR; hops++) {
      if (hops > MAX_SECTORS) return null;
      const sec = sectorAt(difat);
      if (!sec) return null;
      for (let i = 0; i < perSector - 1 && fatSectors.length < fatSectorCount; i++) {
        const s = sec.readUInt32LE(i * 4);
        if (s !== FREE_SECTOR) fatSectors.push(s);
      }
      difat = sec.readUInt32LE((perSector - 1) * 4);
    }
    if (fatSectors.length < fatSectorCount) return null;

    const fat: number[] = [];
    for (const s of fatSectors) {
      const sec = sectorAt(s);
      if (!sec) return null;
      for (let i = 0; i < perSector; i++) fat.push(sec.readUInt32LE(i * 4));
    }

    const names: string[] = [];
    let sector = firstDirSector;
    const seen = new Set<number>();
    while (sector !== END_OF_CHAIN) {
      if (seen.has(sector) || sector >= fat.length) return null; // a loop or a sector outside the table
      seen.add(sector);
      const sec = sectorAt(sector);
      if (!sec) return null;
      for (let off = 0; off + 128 <= sectorSize; off += 128) {
        const type = sec[off + 66];
        if (type === 0) continue; // unused entry
        const len = sec.readUInt16LE(off + 64);
        if (len < 2 || len > 64 || len % 2 !== 0) return null;
        names.push(sec.subarray(off, off + len - 2).toString('utf16le'));
      }
      sector = fat[sector];
    }
    return names;
  } catch {
    return null;
  }
}

/** True only for a compound file that is exactly the wrapper Word writes around a password-protected .docx. */
export function isEncryptedOoxmlPackage(buf: Buffer): boolean {
  const names = listCfbEntryNames(buf);
  if (!names || names.length === 0) return false;
  const lower = names.map((n) => n.toLowerCase());
  if (!lower.includes('encryptioninfo') || !lower.includes('encryptedpackage')) return false;
  return lower.every((n) => ENCRYPTED_PACKAGE_NAMES.has(n));
}
