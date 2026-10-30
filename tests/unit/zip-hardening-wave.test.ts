import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import crypto from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  createEmptyDocument,
  readDocxFileBuffer,
  ZIP_FIXED_EPOCH_MS,
  ZIP_BOMB_MAX_UNCOMPRESSED_BYTES,
  ZIP_BOMB_MAX_RATIO,
  ZIP_CEILING_MAX_RATIO_ENV,
  ZIP_CEILING_MAX_UNCOMPRESSED_ENV,
  resolveDeterministicZipDate,
  resolveMaxRatio,
  resolveMaxUncompressedBytes,
  resolveZipBombCeilings,
  checkZipBombCeilings,
  toDosTimeUtc,
  ZipLike,
} from '@attentionform/astra-engine';
import { run as runCli, EXIT_OK, EXIT_FAIL } from '../../packages/cli/src/cli';

/**
 * Binary helper to patch Central Directory headers in a zip archive without inflating.
 */
function patchCentralDirectory(
  buffer: Buffer,
  patchFn: (entry: { name: string; compressedSize: number; uncompressedSize: number }) => {
    compressedSize?: number;
    uncompressedSize?: number;
  } | void
): Buffer {
  const copy = Buffer.from(buffer);
  let offset = 0;
  const cdSig = Buffer.from([0x50, 0x4b, 0x01, 0x02]);
  while ((offset = copy.indexOf(cdSig, offset)) !== -1) {
    const compSize = copy.readUInt32LE(offset + 20);
    const uncompSize = copy.readUInt32LE(offset + 24);
    const nameLen = copy.readUInt16LE(offset + 28);
    const name = copy.toString('utf8', offset + 46, offset + 46 + nameLen);
    const res = patchFn({ name, compressedSize: compSize, uncompressedSize: uncompSize });
    if (res) {
      if (typeof res.compressedSize === 'number') {
        copy.writeUInt32LE(res.compressedSize, offset + 20);
      }
      if (typeof res.uncompressedSize === 'number') {
        copy.writeUInt32LE(res.uncompressedSize, offset + 24);
      }
    }
    offset += 4;
  }
  return copy;
}

function sha256(buf: Buffer): string {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

describe('Wave 1F: Deterministic zip timestamps (P0-8) & Zip bomb ceilings (T-2)', () => {
  const originalEnv = { ...process.env };

  beforeEach(() => {
    delete process.env.SOURCE_DATE_EPOCH;
    // Pinned output is opt-in; these tests are about what it guarantees.
    process.env.ASTRA_REPRODUCIBLE = '1';
    delete process.env.ASTRA_NO_REPRODUCIBLE;
  });

  afterEach(() => {
    process.env = { ...originalEnv };
  });

  // Test 1: Byte-identical output across distinct saves at different wall-clock times
  describe('Test 1: Byte-identical output across distinct saves at different wall-clock times', () => {
    it('produces identical SHA-256 hashes and byte-identical buffers across saves at different times', async () => {
      const docx1 = createBlankDocxBuffer('blank');
      const hash1 = sha256(docx1);

      // Advance wall-clock time
      await new Promise((resolve) => setTimeout(resolve, 50));

      const docx2 = createBlankDocxBuffer('blank');
      const hash2 = sha256(docx2);

      expect(hash1).toBe(hash2);
      expect(docx1.equals(docx2)).toBe(true);

      // Verify successive saves from the same parser instance at different wall-clock times
      const parser = new DocxParser(docx1);
      const saveA = parser.rezip();
      await new Promise((resolve) => setTimeout(resolve, 50));
      const saveB = parser.rezip();

      expect(sha256(saveA)).toBe(sha256(saveB));
      expect(saveA.equals(saveB)).toBe(true);
    });

    it('produces identical SHA-256 hashes for distinct parsers with identical mutations', async () => {
      const baseBuffer = createBlankDocxBuffer('blank');

      const p1 = new DocxParser(baseBuffer);
      const summary1 = p1.getStructuralSummary();
      p1.updateNode(summary1.nodes[0].id, '<w:p><w:r><w:t>Deterministic Text</w:t></w:r></w:p>');
      const out1 = p1.rezip();

      await new Promise((resolve) => setTimeout(resolve, 50));

      const p2 = new DocxParser(baseBuffer);
      const summary2 = p2.getStructuralSummary();
      p2.updateNode(summary2.nodes[0].id, '<w:p><w:r><w:t>Deterministic Text</w:t></w:r></w:p>');
      const out2 = p2.rezip();

      expect(sha256(out1)).toBe(sha256(out2));
      expect(out1.equals(out2)).toBe(true);
    });

    it('produces byte-identical SHA-256 hashes across different timezones (America/New_York, UTC, Asia/Tokyo, Europe/London)', () => {
      const timezones = ['America/New_York', 'UTC', 'Asia/Tokyo', 'Europe/London'];
      const hashes = new Set<string>();
      const initialTz = process.env.TZ;
      try {
        for (const tz of timezones) {
          process.env.TZ = tz;
          const buf = createBlankDocxBuffer('blank');
          hashes.add(sha256(buf));
        }
        expect(hashes.size).toBe(1);
      } finally {
        if (initialTz !== undefined) {
          process.env.TZ = initialTz;
        } else {
          delete process.env.TZ;
        }
      }
    });

    it('produces byte-identical SHA-256 hashes with custom SOURCE_DATE_EPOCH across different timezones', () => {
      const timezones = ['America/New_York', 'UTC', 'Asia/Tokyo', 'Europe/London'];
      const hashes = new Set<string>();
      const initialTz = process.env.TZ;
      const initialEpoch = process.env.SOURCE_DATE_EPOCH;
      process.env.SOURCE_DATE_EPOCH = '1700000000';
      try {
        for (const tz of timezones) {
          process.env.TZ = tz;
          const parser = DocxParser.createNewDocument('blank');
          const buf = parser.rezip();
          hashes.add(sha256(buf));
        }
        expect(hashes.size).toBe(1);
      } finally {
        if (initialEpoch !== undefined) {
          process.env.SOURCE_DATE_EPOCH = initialEpoch;
        } else {
          delete process.env.SOURCE_DATE_EPOCH;
        }
        if (initialTz !== undefined) {
          process.env.TZ = initialTz;
        } else {
          delete process.env.TZ;
        }
      }
    });
  });

  // Test 2: Entry timestamps clamped to 1980-01-01 UTC (DOS earliest representable time)
  describe('Test 2: Entry timestamps clamped to 1980-01-01 UTC', () => {
    it('clamps all zip entries to 1980-01-01T00:00:00Z on rezip', () => {
      const docx = createBlankDocxBuffer('blank');
      const zip = new AdmZip(docx);
      const entries = zip.getEntries();

      expect(entries.length).toBeGreaterThan(0);
      const expectedEpoch = new Date(ZIP_FIXED_EPOCH_MS);
      const expectedDosTime = toDosTimeUtc(expectedEpoch);

      for (const entry of entries) {
        expect((entry.header as any).timeval).toBe(expectedDosTime);
        expect(entry.header.time.getFullYear()).toBe(1980);
        expect(entry.header.time.getMonth()).toBe(0);
        expect(entry.header.time.getDate()).toBe(1);
        expect(entry.header.time.getHours()).toBe(0);
        expect(entry.header.time.getMinutes()).toBe(0);
        expect(entry.header.time.getSeconds()).toBe(0);
      }
    });

    it('clamps all zip entries in DocxParser.createNewDocument and createEmptyDocument', () => {
      const expectedDosTime = toDosTimeUtc(new Date(ZIP_FIXED_EPOCH_MS));

      const parser = DocxParser.createNewDocument('blank');
      const buf = parser.rezip();
      const zip = new AdmZip(buf);
      for (const entry of zip.getEntries()) {
        expect((entry.header as any).timeval).toBe(expectedDosTime);
      }

      const emptyParser = createEmptyDocument('blank');
      const emptyBuf = emptyParser.rezip();
      const emptyZip = new AdmZip(emptyBuf);
      for (const entry of emptyZip.getEntries()) {
        expect((entry.header as any).timeval).toBe(expectedDosTime);
      }
    });
  });

  // Test 3: SOURCE_DATE_EPOCH support and pre-1980 clamping
  describe('Test 3: SOURCE_DATE_EPOCH support and pre-1980 clamping', () => {
    it('uses fallback 1980-01-01 UTC when SOURCE_DATE_EPOCH is unset, empty, or whitespace', () => {
      expect(resolveDeterministicZipDate({})).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
      expect(resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: '' })).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
      expect(resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: '   ' })).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
    });

    it('falls back to 1980-01-01 UTC when SOURCE_DATE_EPOCH is non-numeric or negative', () => {
      expect(resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: 'not-a-number' })).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
      expect(resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: '-100' })).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
    });

    it('clamps up to 1980-01-01 UTC for pre-1980 timestamps', () => {
      // Unix epoch: 1970-01-01
      expect(resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: '0' })).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
      // 1979-12-31T23:59:59Z (one second before DOS epoch)
      const justBefore1980 = Math.floor(ZIP_FIXED_EPOCH_MS / 1000) - 1;
      expect(resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: String(justBefore1980) })).toEqual(new Date(ZIP_FIXED_EPOCH_MS));
    });

    it('honors valid post-1980 SOURCE_DATE_EPOCH values', () => {
      const timestampSecs = 1700000000; // 2023-11-14T22:13:20.000Z
      const resolved = resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: String(timestampSecs) });
      expect(resolved.getTime()).toBe(timestampSecs * 1000);
      expect(resolved.toISOString()).toBe('2023-11-14T22:13:20.000Z');

      // Floored decimal seconds
      const resolvedDec = resolveDeterministicZipDate({ SOURCE_DATE_EPOCH: '1700000000.850' });
      expect(resolvedDec.getTime()).toBe(timestampSecs * 1000);
    });

    it('applies SOURCE_DATE_EPOCH through rezip()', () => {
      process.env.SOURCE_DATE_EPOCH = '1700000000';
      const docx = createBlankDocxBuffer('blank');
      const zip = new AdmZip(docx);
      const entries = zip.getEntries();
      const expectedEpoch = new Date(1700000000 * 1000);
      const expectedDosTime = toDosTimeUtc(expectedEpoch);
      for (const entry of entries) {
        expect((entry.header as any).timeval).toBe(expectedDosTime);
        expect(entry.header.time.getFullYear()).toBe(expectedEpoch.getUTCFullYear());
        expect(entry.header.time.getMonth()).toBe(expectedEpoch.getUTCMonth());
        expect(entry.header.time.getDate()).toBe(expectedEpoch.getUTCDate());
        expect(entry.header.time.getHours()).toBe(expectedEpoch.getUTCHours());
        expect(entry.header.time.getMinutes()).toBe(expectedEpoch.getUTCMinutes());
        expect(entry.header.time.getSeconds()).toBe(expectedEpoch.getUTCSeconds());
      }
    });
  });

  // Test 4: Zip bomb total uncompressed ceiling (> 200 MiB) fails closed with [E_LIMIT] without inflating
  describe('Test 4: Zip bomb total uncompressed ceiling (> 200 MiB) fails closed with [E_LIMIT]', () => {
    it('checkZipBombCeilings throws [E_LIMIT] when total declared uncompressed bytes exceed default 200 MiB', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'large-entry-1.xml',
            header: {
              size: 150 * 1024 * 1024,
              compressedSize: 5 * 1024 * 1024,
              time: new Date(),
            },
          },
          {
            entryName: 'large-entry-2.xml',
            header: {
              size: 60 * 1024 * 1024,
              compressedSize: 2 * 1024 * 1024,
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: total uncompressed 220200960 bytes exceeds 209715200 bytes/
      );
    });

    it('enforces custom maxUncompressedBytes option', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'entry.xml',
            header: {
              size: 10000,
              compressedSize: 200,
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip, { maxUncompressedBytes: 5000 })).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: total uncompressed 10000 bytes exceeds 5000 bytes/
      );
    });

    it('DocxParser fails closed on crafted zip bomb declaring > 200 MiB without inflating', () => {
      const validDocx = createBlankDocxBuffer('blank');

      // Craft a zip whose Central Directory declares uncompressedSize > 200 MiB
      const crafted = patchCentralDirectory(validDocx, (entry) => {
        if (entry.name === 'word/document.xml') {
          return { uncompressedSize: 210 * 1024 * 1024, compressedSize: 5 * 1024 * 1024 };
        }
      });

      // Loading via DocxParser must throw [E_LIMIT] synchronously BEFORE inflating
      expect(() => new DocxParser(crafted)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: total uncompressed .* exceeds 209715200 bytes/
      );
    });
  });

  // Test 5: Zip bomb ratio ceiling (> 100:1) on >= 1KB entry fails closed with [E_LIMIT]
  describe('Test 5: Zip bomb ratio ceiling (> 100:1) on >= 1KB entry fails closed with [E_LIMIT]', () => {
    it('skips entry ratio check for entries smaller than ZIP_BOMB_RATIO_MIN_ENTRY_BYTES (1024 bytes)', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'tiny.xml',
            header: {
              size: 1000, // < 1024 bytes
              compressedSize: 1, // ratio 1000:1 would exceed 100:1 if checked
              time: new Date(),
            },
          },
        ],
      };

      // Does not throw because entry < 1KB
      expect(() => checkZipBombCeilings(mockZip)).not.toThrow();
    });

    it('throws [E_LIMIT] when entry size >= 1024 and ratio > 100:1', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'word/document.xml',
            header: {
              size: 10240, // 10 KB
              compressedSize: 50, // ratio 204.8:1 > 100:1
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'word\/document\.xml' ratio 204.8:1 exceeds 100:1 \(size 10240, compressed 50\)/
      );
    });

    it('DocxParser fails closed when an entry exhibits ratio > 100:1', () => {
      const validDocx = createBlankDocxBuffer('blank');

      // Patch styles.xml in CD to claim ratio > 100:1
      const crafted = patchCentralDirectory(validDocx, (entry) => {
        if (entry.name === 'word/styles.xml') {
          return { uncompressedSize: 5000, compressedSize: 20 }; // 250:1 ratio
        }
      });

      expect(() => new DocxParser(crafted)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'word\/styles\.xml' ratio 250\.0:1 exceeds 100:1/
      );
    });
  });

  // Test 6: Overall wire-to-uncompressed ratio (> 100:1) fails closed with [E_LIMIT]
  describe('Test 6: Overall wire-to-uncompressed ratio (> 100:1) fails closed with [E_LIMIT]', () => {
    it('throws [E_LIMIT] when overall ratio > 100:1 even if individual entries skip entry ratio check', () => {
      // 20 entries of 500 bytes = 10,000 bytes uncompressed (each < 1024 bytes, so entry ratio skipped)
      const entries = Array.from({ length: 20 }, (_, i) => ({
        entryName: `part${i}.xml`,
        header: {
          size: 500,
          compressedSize: 400,
          time: new Date(),
        },
      }));

      const mockZip: ZipLike = { getEntries: () => entries };

      // Wire is 80 bytes -> overall ratio = 10000 / 80 = 125:1 > 100:1
      expect(() => checkZipBombCeilings(mockZip, { sourceCompressedBytes: 80 })).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: overall ratio 125\.0:1 exceeds 100:1 \(total 10000, wire 80\)/
      );
    });

    it('DocxParser fails closed when overall wire-to-uncompressed ratio > 100:1', () => {
      const validDocx = createBlankDocxBuffer('blank');

      // Patch entries so individual ratios are ~83:1 (< 100:1 ceiling), but total uncompressed is 250,000 bytes.
      // With wire size ~2KB, overall ratio is ~125:1 (> 100:1 ceiling).
      const crafted = patchCentralDirectory(validDocx, () => {
        return { uncompressedSize: 50000, compressedSize: 600 };
      });

      expect(() => new DocxParser(crafted)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: overall ratio .* exceeds 100:1/
      );
    });
  });

  // Test 7: Standard valid docx packages load and save cleanly without false positives
  describe('Test 7: Standard valid docx packages load and save cleanly without false positives', () => {
    it('loads and rezips standard blank document without false positives', () => {
      const buf = createBlankDocxBuffer('blank');
      const parser = new DocxParser(buf);
      const summary = parser.getStructuralSummary();
      expect(summary.nodes.length).toBeGreaterThan(0);

      const rezipped = parser.rezip();
      expect(rezipped.length).toBeGreaterThan(0);

      // Re-load to ensure roundtrip integrity
      const roundtrip = new DocxParser(rezipped);
      expect(roundtrip.getStructuralSummary().nodes.length).toBe(summary.nodes.length);
    });

    it('loads and rezips built-in templates cleanly', () => {
      const templates: Array<'blank' | 'novel' | 'financial' | 'report'> = ['blank', 'novel', 'financial', 'report'];
      for (const t of templates) {
        const p = DocxParser.createNewDocument(t);
        const rezipped = p.rezip();
        expect(rezipped.length).toBeGreaterThan(0);

        const loaded = new DocxParser(rezipped);
        expect(loaded.getStructuralSummary().nodes.length).toBeGreaterThan(0);
      }
    });

    it('createEmptyDocument static method and export function behave identically', () => {
      const p1 = DocxParser.createEmptyDocument('blank');
      const p2 = createEmptyDocument('blank');

      const buf1 = p1.rezip();
      const buf2 = p2.rezip();

      expect(sha256(buf1)).toBe(sha256(buf2));
    });

    it('supports AST node mutation and saves cleanly', () => {
      const buf = createBlankDocxBuffer('blank');
      const parser = new DocxParser(buf);
      const summary = parser.getStructuralSummary();
      const node0 = summary.nodes[0];

      parser.updateNode(node0.id, '<w:p><w:r><w:t>Updated Content Cleanly</w:t></w:r></w:p>');
      const updatedBuf = parser.rezip();

      const parserReloaded = new DocxParser(updatedBuf);
      const nodeText = parserReloaded.readNode(node0.id);
      expect(nodeText).toContain('Updated Content Cleanly');
    });
  });

  // Test 8: Zero-size payload bypass and invalid/negative size validation (fails closed)
  describe('Test 8: Zero-size payload bypass and invalid/negative size validation (fails closed)', () => {
    it('throws [E_LIMIT] when an entry declares uncompressed size 0 with non-zero compressed payload', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'zero-uncompressed-bomb.xml',
            header: {
              size: 0,
              compressedSize: 500,
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'zero-uncompressed-bomb\.xml' declared size 0 with non-zero compressed payload \(500 bytes\)/
      );
    });

    it('DocxParser fails closed when an entry declares uncompressed size 0 with non-zero compressed payload', () => {
      const validDocx = createBlankDocxBuffer('blank');
      const crafted = patchCentralDirectory(validDocx, (entry) => {
        if (entry.name === 'word/document.xml') {
          return { uncompressedSize: 0, compressedSize: 1024 };
        }
      });

      expect(() => new DocxParser(crafted)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'word\/document\.xml' declared size 0 with non-zero compressed payload \(1024 bytes\)/
      );
    });

    it('permits entries where both uncompressed size and compressed size are 0', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'empty-file.txt',
            header: {
              size: 0,
              compressedSize: 0,
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip)).not.toThrow();
    });

    it('throws [E_LIMIT] when entry uncompressed size is negative', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'negative-size.xml',
            header: {
              size: -100,
              compressedSize: 50,
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'negative-size\.xml' has invalid declared size/
      );
    });

    it('throws [E_LIMIT] when entry compressedSize is negative', () => {
      const mockZip: ZipLike = {
        getEntries: () => [
          {
            entryName: 'negative-comp.xml',
            header: {
              size: 100,
              compressedSize: -1,
              time: new Date(),
            },
          },
        ],
      };

      expect(() => checkZipBombCeilings(mockZip)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'negative-comp\.xml' has invalid declared size/
      );
    });

    it('throws [E_LIMIT] when entry size is non-finite (NaN or Infinity)', () => {
      const mockZipNaN: ZipLike = {
        getEntries: () => [
          {
            entryName: 'nan-size.xml',
            header: {
              size: NaN,
              compressedSize: 10,
              time: new Date(),
            },
          },
        ],
      };
      expect(() => checkZipBombCeilings(mockZipNaN)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'nan-size\.xml' has invalid declared size/
      );

      const mockZipInf: ZipLike = {
        getEntries: () => [
          {
            entryName: 'inf-size.xml',
            header: {
              size: Infinity,
              compressedSize: 10,
              time: new Date(),
            },
          },
        ],
      };
      expect(() => checkZipBombCeilings(mockZipInf)).toThrowError(
        /\[E_LIMIT\] zip bomb ceiling: entry 'inf-size\.xml' has invalid declared size/
      );
    });
  });
});

describe('R-8: tunable zip-bomb ceilings (env + flags) and single-prefix E_LIMIT', () => {
  const RATIO_ENV = ZIP_CEILING_MAX_RATIO_ENV;
  const BYTES_ENV = ZIP_CEILING_MAX_UNCOMPRESSED_ENV;
  let savedRatio: string | undefined;
  let savedBytes: string | undefined;

  beforeEach(() => {
    savedRatio = process.env[RATIO_ENV];
    savedBytes = process.env[BYTES_ENV];
    delete process.env[RATIO_ENV];
    delete process.env[BYTES_ENV];
  });

  afterEach(() => {
    if (savedRatio === undefined) delete process.env[RATIO_ENV];
    else process.env[RATIO_ENV] = savedRatio;
    if (savedBytes === undefined) delete process.env[BYTES_ENV];
    else process.env[BYTES_ENV] = savedBytes;
    vi.restoreAllMocks();
  });

  // ~44 KB highly-repetitive entry: 45056 declared bytes over 400 compressed
  // bytes = 112.6:1, over the 100:1 default but lawful content.
  function repeatEntryZip(): ZipLike {
    return {
      getEntries: () => [
        {
          entryName: 'word/document.xml',
          header: { size: 45056, compressedSize: 400, time: new Date() },
        },
      ],
    };
  }

  function craftRepeatDocx(): Buffer {
    // Genuinely high-ratio, fully lawful package: minimal document.xml with a
    // 43 KB single-run paragraph (43,387 declared bytes over 343 compressed
    // bytes = 126.5:1). Headers are truthful, so the payload inflates fine
    // once the ceiling is raised — no header lying involved.
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const run = 'A'.repeat(43000);
    const minDoc =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">${run}</w:t></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>';
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.deleteFile('word/document.xml');
    zip.addFile('word/document.xml', Buffer.from(minDoc, 'utf8'));
    return zip.toBuffer();
  }

  function captureConsole() {
    const out: string[] = [];
    const err: string[] = [];
    vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => {
      out.push(a.map(String).join(' '));
    });
    vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => {
      err.push(a.map(String).join(' '));
    });
    return { out, err };
  }

  function writeTmp(buf: Buffer, suffix: string): string {
    const p = path.join(os.tmpdir(), `astra-r8-${process.pid}-${Date.now()}-${suffix}.docx`);
    fs.writeFileSync(p, buf);
    return p;
  }

  it('defaults are unchanged (100:1, 200 MiB) and invalid env falls back closed', () => {
    expect(ZIP_BOMB_MAX_RATIO).toBe(100);
    expect(ZIP_BOMB_MAX_UNCOMPRESSED_BYTES).toBe(200 * 1024 * 1024);
    expect(resolveZipBombCeilings({})).toEqual({
      maxRatio: 100,
      maxUncompressedBytes: 200 * 1024 * 1024,
    });
    for (const bad of ['', '   ', 'abc', '-5', '0', 'NaN', 'Infinity']) {
      expect(resolveMaxRatio({ [RATIO_ENV]: bad })).toBe(100);
      expect(resolveMaxUncompressedBytes({ [BYTES_ENV]: bad })).toBe(200 * 1024 * 1024);
    }
  });

  it('env overrides resolve (ASTRA_MAX_RATIO / ASTRA_MAX_UNCOMPRESSED_BYTES)', () => {
    expect(resolveMaxRatio({ [RATIO_ENV]: '200' })).toBe(200);
    expect(resolveMaxRatio({ [RATIO_ENV]: ' 150.5 ' })).toBe(150.5);
    expect(resolveMaxUncompressedBytes({ [BYTES_ENV]: '314572800' })).toBe(314572800);
    expect(resolveZipBombCeilings({ [RATIO_ENV]: '250', [BYTES_ENV]: '1000' })).toEqual({
      maxRatio: 250,
      maxUncompressedBytes: 1000,
    });
  });

  it('legit 44 KB repeat entry is rejected at default but passes with raised knob', () => {
    expect(() => checkZipBombCeilings(repeatEntryZip())).toThrowError(
      /\[E_LIMIT\] zip bomb ceiling: entry 'word\/document\.xml' ratio 112\.6:1 exceeds 100:1/
    );
    expect(() => checkZipBombCeilings(repeatEntryZip(), { maxRatio: 200 })).not.toThrow();
    process.env[RATIO_ENV] = '200';
    expect(() => checkZipBombCeilings(repeatEntryZip())).not.toThrow();
  });

  it('ratio remedy names the knob with observed vs allowed values', () => {
    try {
      checkZipBombCeilings(repeatEntryZip());
      expect.unreachable('should have thrown');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      expect(msg).toMatch(/Remedy: raise ASTRA_MAX_RATIO \/ --max-ratio above 112\.6 \(current ceiling 100\)/);
    }
  });

  it('total-bytes remedy names its knob with observed vs allowed values', () => {
    const mockZip: ZipLike = {
      getEntries: () => [
        { entryName: 'entry.xml', header: { size: 10000, compressedSize: 200, time: new Date() } },
      ],
    };
    try {
      checkZipBombCeilings(mockZip, { maxUncompressedBytes: 5000 });
      expect.unreachable('should have thrown');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      expect(msg).toMatch(
        /Remedy: raise ASTRA_MAX_UNCOMPRESSED_BYTES \/ --max-uncompressed-bytes above 10000 \(current ceiling 5000\)/
      );
    }
  });

  it('real DocxParser load honors ASTRA_MAX_RATIO env (reject default, load raised)', () => {
    const crafted = craftRepeatDocx();
    expect(() => new DocxParser(crafted)).toThrowError(
      /\[E_LIMIT\] zip bomb ceiling: entry 'word\/document\.xml' ratio 126\.5:1 exceeds 100:1/
    );
    process.env[RATIO_ENV] = '200';
    const parser = new DocxParser(crafted);
    expect(parser.getStructuralSummary().nodes.length).toBeGreaterThan(0);
  });

  it('CLI validate renders E_LIMIT with a single prefix and actionable remedy', async () => {
    const file = writeTmp(craftRepeatDocx(), 'validate');
    try {
      const cap = captureConsole();
      const code = await runCli(['validate', file]);
      expect(code).toBe(EXIT_FAIL);
      const stderr = cap.err.join('\n');
      expect((stderr.match(/\[E_LIMIT\]/g) ?? []).length).toBe(1);
      expect(stderr).not.toContain('[E_VALIDATION]');
      expect(stderr).not.toContain('[E_LIMIT][E_LIMIT]');
      expect(stderr).toMatch(/Remedy:.*ASTRA_MAX_RATIO/);
      expect(stderr).toMatch(/126\.5/);
    } finally {
      fs.unlinkSync(file);
    }
  });

  it('CLI validate --max-ratio lets the legit repeat doc through; bad values are usage errors', async () => {
    const file = writeTmp(craftRepeatDocx(), 'flag');
    try {
      const cap = captureConsole();
      const code = await runCli(['validate', file, '--max-ratio', '200']);
      expect(code).toBe(EXIT_OK);
      expect(cap.out.join('\n')).toContain(`VALID ${file}`);

      const cap2 = captureConsole();
      const bad = await runCli(['validate', file, '--max-ratio', 'banana']);
      expect(bad).toBe(2);
      expect(cap2.err.join('\n')).toMatch(/--max-ratio must be/);
    } finally {
      fs.unlinkSync(file);
    }
  });

  it('CLI inspect honors --max-ratio (reject default, print raised)', async () => {
    const file = writeTmp(craftRepeatDocx(), 'inspect');
    try {
      const cap = captureConsole();
      expect(await runCli(['inspect', file])).toBe(EXIT_FAIL);
      expect((cap.err.join('\n').match(/\[E_LIMIT\]/g) ?? []).length).toBe(1);

      const cap2 = captureConsole();
      expect(await runCli(['inspect', file, '--max-ratio=200'])).toBe(EXIT_OK);
      expect(cap2.out.join('\n')).toMatch(/rows=\d+/);
    } finally {
      fs.unlinkSync(file);
    }
  });

  it('CLI batch honors --max-ratio and keeps E_LIMIT single-prefix at default', async () => {
    const file = writeTmp(craftRepeatDocx(), 'batch-in');
    const outFile = path.join(os.tmpdir(), `astra-r8-${process.pid}-${Date.now()}-batch-out.docx`);
    try {
      process.env[RATIO_ENV] = '200';
      const target = new DocxParser(fs.readFileSync(file)).getStructuralSummary().nodes[0].id;
      delete process.env[RATIO_ENV];
      const opsPath = path.join(os.tmpdir(), `astra-r8-${process.pid}-${Date.now()}-ops.json`);
      fs.writeFileSync(opsPath, JSON.stringify([{ action: 'replace', target, text: 'R8 batch ok' }]));
      try {
        const cap = captureConsole();
        expect(await runCli(['batch', file, opsPath, '--out', outFile])).toBe(EXIT_FAIL);
        expect((cap.err.join('\n').match(/\[E_LIMIT\]/g) ?? []).length).toBe(1);

        const cap2 = captureConsole();
        expect(await runCli(['batch', file, opsPath, '--out', outFile, '--max-ratio', '200'])).toBe(EXIT_OK);
        expect(cap2.out.join('\n')).toContain('BATCH OK');
        expect(fs.existsSync(outFile)).toBe(true);
      } finally {
        fs.unlinkSync(opsPath);
      }
    } finally {
      fs.unlinkSync(file);
      try {
        fs.unlinkSync(outFile);
      } catch {
        /* not written on failure paths */
      }
    }
  });

  it('tunables are documented in --help for validate/inspect/batch', async () => {
    for (const cmd of ['validate', 'inspect', 'batch'] as const) {
      const cap = captureConsole();
      const code = await runCli([cmd, '--help']);
      expect(code).toBe(EXIT_OK);
      const text = cap.out.join('\n');
      expect(text).toContain('--max-ratio');
      expect(text).toContain('--max-uncompressed-bytes');
      expect(text).toContain('ASTRA_MAX_RATIO');
      expect(text).toContain('ASTRA_MAX_UNCOMPRESSED_BYTES');
    }
  });

  it('pre-read gate honors raised env knob (reject low, pass raised)', () => {
    const p = path.join(os.tmpdir(), `astra-r8-${process.pid}-${Date.now()}-preread.bin`);
    fs.writeFileSync(p, Buffer.alloc(100, 0x41));
    try {
      process.env[BYTES_ENV] = '50';
      expect(() => readDocxFileBuffer(p)).toThrowError(/\[E_LIMIT\].*pre-read ceiling 50 bytes/);
      process.env[BYTES_ENV] = '200';
      expect(() => readDocxFileBuffer(p)).not.toThrow();
      expect(readDocxFileBuffer(p).length).toBe(100);
    } finally {
      fs.unlinkSync(p);
    }
  });

  it('pre-read gate falls back to default on invalid env (fail-closed)', () => {
    const p = path.join(os.tmpdir(), `astra-r8-${process.pid}-${Date.now()}-preread-invalid.bin`);
    fs.writeFileSync(p, Buffer.alloc(100, 0x41));
    try {
      for (const bad of ['banana', '0', '-5', '', 'Infinity', 'NaN']) {
        process.env[BYTES_ENV] = bad;
        expect(readDocxFileBuffer(p).length).toBe(100);
      }
    } finally {
      fs.unlinkSync(p);
    }
  });

  it('pre-read remedy names the knob with observed vs ceiling values', () => {
    const p = path.join(os.tmpdir(), `astra-r8-${process.pid}-${Date.now()}-preread-remedy.bin`);
    fs.writeFileSync(p, Buffer.alloc(100, 0x41));
    try {
      process.env[BYTES_ENV] = '50';
      try {
        readDocxFileBuffer(p);
        expect.unreachable('should have thrown');
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        expect(msg).toMatch(/Remedy: raise ASTRA_MAX_UNCOMPRESSED_BYTES \/ --max-uncompressed-bytes above 100 \(current ceiling 50\)/);
      }
    } finally {
      fs.unlinkSync(p);
    }
  });

  it('CLI E_LIMIT line has no doubled period before Remedy', async () => {
    const file = writeTmp(craftRepeatDocx(), 'double-period');
    try {
      const cap = captureConsole();
      const code = await runCli(['validate', file]);
      expect(code).toBe(EXIT_FAIL);
      const stderr = cap.err.join('\n');
      expect(stderr).not.toContain('.. Remedy:');
      expect(stderr).toMatch(/\. Remedy: /);
      expect((stderr.match(/Remedy:/g) ?? []).length).toBe(1);
    } finally {
      fs.unlinkSync(file);
    }
  });
});
