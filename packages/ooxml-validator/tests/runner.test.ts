// Adapted from xarsh/ooxml-validator (MIT). See LICENSE-MIT-xarsh.

import { describe, it, expect } from 'vitest';
import { copyFileSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { isValid, OFFICE_VERSIONS, validateFile } from '../src/index.js';
import {
  MANY_ERRORS_DOCX,
  OFFICE2010_ONLY_DOCX,
  VALID_PPTX,
  VALID_DOCX,
  INVALID_SCHEMA_DOCX,
  CORRUPT_ZIP_DOCX,
} from './helpers.js';

describe('validateFile and isValid', () => {
  it('validateFile returns ok=true for a valid pptx', async () => {
    const result = await validateFile(VALID_PPTX);
    expect(result.ok).toBe(true);
    expect(result.errors.length).toBe(0);
    expect(result.file).toBe(VALID_PPTX);
  });

  it('isValid returns true for a valid pptx', async () => {
    expect(await isValid(VALID_PPTX)).toBe(true);
  });

  it('validateFile returns ok=true for a valid docx', async () => {
    const result = await validateFile(VALID_DOCX);
    expect(result.ok).toBe(true);
    expect(result.errors.length).toBe(0);
  });

  it('validateFile returns ok=false for invalid schema docx', async () => {
    const result = await validateFile(INVALID_SCHEMA_DOCX);
    expect(result.ok).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.errors[0].errorType).toBe('Schema');
  });

  it('validateFile returns ok=false for corrupt zip docx', async () => {
    const result = await validateFile(CORRUPT_ZIP_DOCX);
    expect(result.ok).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.errors[0].errorType).toBe('Exception');
  });

  it('officeVersion selects the schema the file is checked against', async () => {
    expect(await isValid(OFFICE2010_ONLY_DOCX, { officeVersion: 'Office2007' })).toBe(false);
    expect(await isValid(OFFICE2010_ONLY_DOCX, { officeVersion: 'Office2010' })).toBe(true);
    expect(await isValid(OFFICE2010_ONLY_DOCX)).toBe(true); // default is Microsoft365
  });

  it('every documented office version reaches the validator', async () => {
    for (const version of OFFICE_VERSIONS) {
      const valid = await isValid(OFFICE2010_ONLY_DOCX, { officeVersion: version });
      expect(valid).toBe(version !== 'Office2007');
    }
  });

  it('validateFile rejects an officeVersion it does not know', async () => {
    for (const bad of ['office2007', 'Office2018', '1', 'Office2007,Office2010', '']) {
      await expect(() =>
        validateFile(OFFICE2010_ONLY_DOCX, { officeVersion: bad as any }),
      ).rejects.toThrow(/Invalid officeVersion/);
    }
  });

  it('validation errors use the key names the README documents', async () => {
    const result = await validateFile(OFFICE2010_ONLY_DOCX, { officeVersion: 'Office2007' });
    expect(result.ok).toBe(false);
    const [error] = result.errors;
    expect(Object.keys(error).sort()).toEqual([
      'class',
      'description',
      'errorType',
      'id',
      'node',
      'part',
      'path',
      'relatedNode',
      'xPath',
    ]);
    expect(error.errorType).toBe('Schema');
    expect(error.path).toBe('/word/document.xml');
    expect(String(error.xPath)).toMatch(/^\/w:document/);
  });

  it('validateFile reports ok=false for a nonexistent file', async () => {
    const result = await validateFile('/no/such/file.pptx');
    expect(result.ok).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.errors[0].errorType).toBe('Exception');
  });

  it('validateFile reports corruption for a non-zip file', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ooxml-validator-test-'));
    const bogus = join(dir, 'bogus.pptx');
    writeFileSync(bogus, 'this is not a zip file');
    try {
      const result = await validateFile(bogus);
      expect(result.ok).toBe(false);
      expect(result.errors[0].description ?? '').toMatch(/corrupt/i);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('validateFile reports an unsupported extension rather than throwing', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ooxml-validator-test-'));
    const txt = join(dir, 'notes.txt');
    writeFileSync(txt, 'hello');
    try {
      const result = await validateFile(txt);
      expect(result.ok).toBe(false);
      expect(result.errors[0].errorType).toBe('Exception');
      expect(result.errors[0].description ?? '').toMatch(/Unsupported extension/);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('paths with spaces and non-ASCII characters are passed through unmangled', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'ooxml validator test '));
    const file = join(dir, 'ドキュメント 1.pptx');
    copyFileSync(VALID_PPTX, file);
    try {
      const result = await validateFile(file);
      expect(result.ok).toBe(true);
      expect(result.file).toBe(file);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('large validation result JSON beyond pipe buffer size completes cleanly', async () => {
    const result = await validateFile(MANY_ERRORS_DOCX, { officeVersion: 'Office2007' });
    expect(result.ok).toBe(false);
    expect(result.errors.length).toBeGreaterThanOrEqual(400);
  });
});
