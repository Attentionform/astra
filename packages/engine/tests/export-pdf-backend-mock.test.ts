/**
 * PDF export backend — mocked, no live binaries in the default run.
 *
 * Default cases stay spawn-free via the testing override seam
 * (`setPdfBackendBinaryForTesting` short-circuits `getPdfBackendBinary`
 * before any `which`/spawn): filter specs are pure string contracts,
 * a nulled override yields E_DEPENDENCY without spawning, and override
 * caching is proven without probing the system.
 * The explicit-missing-binary path needs real spawns (truthy binary skips
 * the gate, LO spawn fails ENOENT, surfaces E_RENDER) so it lives behind
 * ASTRA_RUN_SLOW=1 and never runs by default.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  convertDocxToPdf,
  getPdfBackendBinary,
  hasPdfBackend,
  pdfFilterSpec,
  resetPdfBackendCache,
  setPdfBackendBinaryForTesting,
} from '../src/export/pdf';
import { DocxParser } from '../src/parser';

afterEach(() => {
  resetPdfBackendCache();
});

// Live spawns run only when explicitly enabled; default runs stay
// mocked (no process spawns) and fast. Enable with ASTRA_RUN_SLOW=1.
const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1';

function blankDocx(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

describe('pdf export backend (mocked)', () => {
  it('filter spec is a pure contract (tagged plain, untagged JSON)', () => {
    expect(pdfFilterSpec(true)).toBe('pdf:writer_pdf_Export');
    expect(pdfFilterSpec(false)).toContain('UseTaggedPDF');
    expect(pdfFilterSpec(false)).toContain('writer_pdf_Export');
  });

  it('hasPdfBackend honors the testing override without probing the system (no spawn)', () => {
    setPdfBackendBinaryForTesting(null);
    try {
      expect(hasPdfBackend()).toBe(false);
      expect(getPdfBackendBinary()).toBeNull();
      setPdfBackendBinaryForTesting('mock-binary');
      expect(hasPdfBackend()).toBe(true);
      expect(getPdfBackendBinary()).toBe('mock-binary');
    } finally {
      resetPdfBackendCache();
    }
  });

  it('getPdfBackendBinary caches the testing override until reset (no spawn)', () => {
    setPdfBackendBinaryForTesting(null);
    try {
      expect(getPdfBackendBinary()).toBeNull();
      expect(hasPdfBackend()).toBe(false);
    } finally {
      resetPdfBackendCache();
    }
    // Reset clears the cache: a fresh override sticks without a system probe.
    setPdfBackendBinaryForTesting('mock-binary-after-reset');
    try {
      expect(getPdfBackendBinary()).toBe('mock-binary-after-reset');
      expect(hasPdfBackend()).toBe(true);
    } finally {
      resetPdfBackendCache();
    }
  });

  it('convert rejects E_DEPENDENCY when backend is nulled (no spawn)', async () => {
    setPdfBackendBinaryForTesting(null);
    try {
      const outPdf = path.join(os.tmpdir(), `astra-pdf-mock-${process.pid}-${Date.now()}.pdf`);
      await expect(convertDocxToPdf(blankDocx(), outPdf)).rejects.toThrow(/E_DEPENDENCY/);
      expect(fs.existsSync(outPdf)).toBe(false);
    } finally {
      resetPdfBackendCache();
    }
  });

  it.runIf(RUN_LIVE)('convert with explicit missing binary surfaces E_RENDER live (spawns; opt-in slow)', async () => {
    setPdfBackendBinaryForTesting('/nonexistent/soffice-astra-mock');
    try {
      const outPdf = path.join(os.tmpdir(), `astra-pdf-mock-explicit-${process.pid}-${Date.now()}.pdf`);
      await expect(convertDocxToPdf(blankDocx(), outPdf)).rejects.toThrow(/E_RENDER/);
      expect(fs.existsSync(outPdf)).toBe(false);
    } finally {
      resetPdfBackendCache();
    }
  }, 15000);
});
