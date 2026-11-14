import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { isEncryptedOoxmlPackage, listCfbEntryNames } from '../../packages/engine/src/visual/cfb';
import { analyseReferences, neutraliseReferences } from '../../packages/engine/src/visual/references';
import { buildCfb } from '../helpers/cfb';

const REAL = fs.readFileSync(path.join(__dirname, '..', 'fixtures', 'encrypted-password-astra-test.docx'));
const WRAPPER = ['EncryptionInfo', 'EncryptedPackage', '\u0006DataSpaces', 'Version', 'DataSpaceMap', 'DataSpaceInfo', 'StrongEncryptionDataSpace', 'TransformInfo', 'StrongEncryptionTransform', '\u0006Primary'];

describe('password-protected packages', () => {
  it('a real one made by Word is recognised, and carries nothing to neutralise', () => {
    expect(listCfbEntryNames(REAL)).toEqual(expect.arrayContaining(['EncryptionInfo', 'EncryptedPackage']));
    expect(isEncryptedOoxmlPackage(REAL)).toBe(true);
    const a = neutraliseReferences(REAL);
    expect(a.docx).toBe(REAL); // the same bytes: nothing was changed
    expect(a.changed).toBe(false);
    expect(a.found).toEqual([]);
    expect(analyseReferences(REAL, { mutate: false }).found).toEqual([]);
  });

  it('the wrapper names, built synthetically, are recognised too (512 and 4096 byte sectors)', () => {
    expect(isEncryptedOoxmlPackage(buildCfb(WRAPPER))).toBe(true);
    expect(isEncryptedOoxmlPackage(buildCfb(WRAPPER, { sectorShift: 12 }))).toBe(true);
    expect(isEncryptedOoxmlPackage(buildCfb(['EncryptionInfo', 'EncryptedPackage']))).toBe(true);
  });

  it.each([
    ['a legacy .doc (WordDocument stream) renamed .docx', ['WordDocument', '1Table', 'SummaryInformation']],
    ['a legacy .doc that also claims to be encrypted', [...WRAPPER, 'WordDocument']],
    ['a spreadsheet (Workbook stream)', ['Workbook', 'SummaryInformation']],
    ['a presentation', ['PowerPoint Document', 'Current User']],
    ['an unknown extra stream next to the wrapper', [...WRAPPER, 'Macros']],
    ['only one of the two encryption streams', ['EncryptionInfo']],
    ['no streams at all', []],
  ])('%s is NOT treated as encrypted: it stays unverifiable and fails closed', (_name, names) => {
    const cfb = buildCfb(names);
    expect(isEncryptedOoxmlPackage(cfb)).toBe(false);
    expect(() => neutraliseReferences(cfb)).toThrow(/E_SECURITY/);
  });

  it('names are compared without case, and a name that only contains a known one does not count', () => {
    expect(isEncryptedOoxmlPackage(buildCfb(['ENCRYPTIONINFO', 'encryptedpackage']))).toBe(true);
    expect(isEncryptedOoxmlPackage(buildCfb(['EncryptionInfo', 'EncryptedPackage', 'EncryptedPackage2']))).toBe(false);
  });

  it('a damaged container is not trusted', () => {
    const good = buildCfb(WRAPPER);
    expect(isEncryptedOoxmlPackage(good.subarray(0, 700))).toBe(false); // directory sector cut off
    expect(isEncryptedOoxmlPackage(buildCfb(WRAPPER, { loopDirectory: true }))).toBe(false); // the directory chain loops
    expect(isEncryptedOoxmlPackage(Buffer.from('D0CF11E0A1B11AE1', 'hex'))).toBe(false); // magic and nothing else
    const badShift = Buffer.from(good);
    badShift.writeUInt16LE(7, 0x1e);
    expect(isEncryptedOoxmlPackage(badShift)).toBe(false);
    const hugeFat = Buffer.from(good);
    hugeFat.writeUInt32LE(0x7fffffff, 0x2c);
    expect(isEncryptedOoxmlPackage(hugeFat)).toBe(false);
    for (let cut = 0; cut < good.length; cut += 37) {
      expect(() => isEncryptedOoxmlPackage(good.subarray(0, cut))).not.toThrow();
    }
  });

  it('random bytes and non-containers never match', () => {
    expect(isEncryptedOoxmlPackage(Buffer.alloc(0))).toBe(false);
    expect(isEncryptedOoxmlPackage(Buffer.from('not a container at all'))).toBe(false);
    expect(isEncryptedOoxmlPackage(Buffer.alloc(4096, 0xd0))).toBe(false);
  });
});
