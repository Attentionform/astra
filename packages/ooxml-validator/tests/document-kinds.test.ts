import { afterAll, describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { createValidator, kindFromPath, validateBuffer, validateFile } from '../src/index.js';
import type { DocumentKind } from '../src/index.js';
import { FIXTURES, VALID_DOCX, VALID_PPTX, INVALID_SCHEMA_DOCX, CORRUPT_ZIP_DOCX, runCli } from './helpers.js';

// Kept out of tests/fixtures: the parity test treats every file there as a golden corpus document.
const FAMILY_FIXTURES = join(FIXTURES, '..', 'fixtures-families');
const VALID_XLSX = join(FAMILY_FIXTURES, 'valid.xlsx');
const INVALID_XLSX = join(FAMILY_FIXTURES, 'invalid-schema.xlsx');
const INVALID_PPTX = join(FAMILY_FIXTURES, 'invalid-schema.pptx');

const FAMILIES: Array<{ kind: DocumentKind; valid: string; invalid: string }> = [
  { kind: 'docx', valid: VALID_DOCX, invalid: INVALID_SCHEMA_DOCX },
  { kind: 'xlsx', valid: VALID_XLSX, invalid: INVALID_XLSX },
  { kind: 'pptx', valid: VALID_PPTX, invalid: INVALID_PPTX },
];

// Main-part content type of the base package for each family, and the variants
// (macro-enabled, template, slide show, add-in) the docs list. Each variant is
// the base package with its main-part content type and file extension changed.
const BASE_TYPE: Record<DocumentKind, { part: string; type: string }> = {
  docx: { part: '/word/document.xml', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' },
  xlsx: { part: '/xl/workbook.xml', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
  pptx: { part: '/ppt/presentation.xml', type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml' },
};

const VARIANTS: Array<{ ext: string; kind: DocumentKind; type: string }> = [
  { ext: '.docm', kind: 'docx', type: 'application/vnd.ms-word.document.macroEnabled.main+xml' },
  { ext: '.dotx', kind: 'docx', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml' },
  { ext: '.dotm', kind: 'docx', type: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml' },
  { ext: '.xlsm', kind: 'xlsx', type: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' },
  { ext: '.xltx', kind: 'xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml' },
  { ext: '.xltm', kind: 'xlsx', type: 'application/vnd.ms-excel.template.macroEnabled.main+xml' },
  { ext: '.xlam', kind: 'xlsx', type: 'application/vnd.ms-excel.addin.macroEnabled.main+xml' },
  { ext: '.pptm', kind: 'pptx', type: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml' },
  { ext: '.potx', kind: 'pptx', type: 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml' },
  { ext: '.potm', kind: 'pptx', type: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml' },
  { ext: '.ppsx', kind: 'pptx', type: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml' },
  { ext: '.ppsm', kind: 'pptx', type: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml' },
];

const dir = mkdtempSync(join(tmpdir(), 'ooxml-kinds-'));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

function makeVariant(basePath: string, kind: DocumentKind, ext: string, type: string): string {
  const zip = new AdmZip(readFileSync(basePath));
  const entry = zip.getEntry('[Content_Types].xml');
  if (!entry) throw new Error('no content types');
  const xml = entry.getData().toString('utf8');
  const { type: baseType } = BASE_TYPE[kind];
  if (!xml.includes(baseType)) throw new Error(`base type for ${kind} not found in ${basePath}`);
  zip.updateFile('[Content_Types].xml', Buffer.from(xml.replace(baseType, type), 'utf8'));
  const out = join(dir, `variant${ext}`);
  writeFileSync(out, zip.toBuffer());
  return out;
}

describe('kindFromPath', () => {
  it('maps every supported extension to its family and is case-insensitive', () => {
    expect(kindFromPath('a.DOCX')).toBe('docx');
    for (const v of VARIANTS) expect(kindFromPath(`file${v.ext}`)).toBe(v.kind);
    expect(kindFromPath('a.xlsx')).toBe('xlsx');
    expect(kindFromPath('a.pptx')).toBe('pptx');
    expect(kindFromPath('a.txt')).toBeUndefined();
    expect(kindFromPath('noextension')).toBeUndefined();
  });
});

describe.each(FAMILIES)('$kind: standalone validateFile', ({ valid, invalid }) => {
  it('passes a valid file and reports the requested version', async () => {
    const res = await validateFile(valid);
    expect(res.ok).toBe(true);
    expect(res.requestedOfficeVersion).toBe('Microsoft365');
  });
  it('fails a schema-invalid file with a Schema error', async () => {
    const res = await validateFile(invalid);
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Schema');
  });
});

describe.each(FAMILIES)('$kind: CLI', ({ valid, invalid }) => {
  it('exits 0 for a valid file and 1 for a schema-invalid file', async () => {
    expect((await runCli([valid])).code).toBe(0);
    const bad = await runCli([invalid]);
    expect(bad.code).toBe(1);
    expect(JSON.parse(bad.stdout).errors[0].errorType).toBe('Schema');
  });
});

describe.each(FAMILIES)('$kind: validateBuffer with kind', ({ kind, valid, invalid }) => {
  it('passes a valid buffer', async () => {
    const res = await validateBuffer(readFileSync(valid), { kind });
    expect(res.ok).toBe(true);
    expect(res.kind).toBe(kind);
    expect(res.office).toBe('Office2019');
    expect(res.requestedOfficeVersion).toBe('Office2019');
  });
  it('fails a schema-invalid buffer with a Schema error', async () => {
    const res = await validateBuffer(readFileSync(invalid), { kind });
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Schema');
  });
});

describe('validateBuffer defaults and failures', () => {
  it('defaults to docx', async () => {
    const res = await validateBuffer(readFileSync(VALID_DOCX));
    expect(res.ok).toBe(true);
    expect(res.kind).toBe('docx');
  });
  it('a valid pptx opened as docx fails naming the requested kind and the native text', async () => {
    const res = await validateBuffer(readFileSync(VALID_PPTX));
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
    expect(res.errors[0].description).toMatch(/^Could not open the package as kind "docx": /);
    expect(res.errors[0].description).not.toMatch(/wrong|corrupt document type/i);
  });
  it('a valid docx opened as xlsx fails the same way', async () => {
    const res = await validateBuffer(readFileSync(VALID_DOCX), { kind: 'xlsx' });
    expect(res.ok).toBe(false);
    expect(res.errors[0].description).toMatch(/^Could not open the package as kind "xlsx": /);
  });
  it('rejects an unknown kind value', async () => {
    await expect(validateBuffer(readFileSync(VALID_DOCX), { kind: 'odt' as DocumentKind })).rejects.toThrow(
      /Invalid kind "odt". Expected one of: docx, pptx, xlsx/,
    );
  });
  it('rejects an unknown officeVersion', async () => {
    await expect(
      validateBuffer(readFileSync(VALID_DOCX), { officeVersion: 'Office1999' as never }),
    ).rejects.toThrow(/Invalid officeVersion "Office1999"/);
  });
  it('an empty buffer fails with an Exception error', async () => {
    const res = await validateBuffer(Buffer.alloc(0));
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
  });
  it('a corrupt package fails with an Exception error', async () => {
    const res = await validateBuffer(readFileSync(CORRUPT_ZIP_DOCX));
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
  });
});

describe('standalone validateFile refuses kind', () => {
  it('throws instead of ignoring it', async () => {
    await expect(validateFile(VALID_DOCX, { kind: 'docx' } as never)).rejects.toThrow(/does not take a 'kind' option/);
  });
});

describe('standalone validateFile failures', () => {
  it('reports an unsupported extension as an Exception error', async () => {
    const txt = join(dir, 'notes.txt');
    writeFileSync(txt, 'hello');
    const res = await validateFile(txt);
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
    expect(res.errors[0].description).toBe('Unsupported extension: .txt');
  });
  it('reports a missing file as an Exception error', async () => {
    const res = await validateFile(join(dir, 'missing.docx'));
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
  });
  it('CLI: unsupported extension exits 1 and missing file exits 1', async () => {
    const txt = join(dir, 'notes2.txt');
    writeFileSync(txt, 'hello');
    expect((await runCli([txt])).code).toBe(1);
    expect((await runCli([join(dir, 'missing.docx')])).code).toBe(1);
  });
});

describe('createValidator (worker) validateFile', () => {
  const validator = createValidator();
  afterAll(() => validator.close());

  it.each(FAMILIES)('infers the family from the extension: valid $kind', async ({ kind, valid }) => {
    const res = await validator.validateFile(valid);
    expect(res.ok).toBe(true);
    expect(res.kind).toBe(kind);
    expect(res.file).toBe(valid);
  });
  it.each(FAMILIES)('fails a schema-invalid $kind with a Schema error', async ({ invalid }) => {
    const res = await validator.validateFile(invalid);
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Schema');
  });
  it('an explicit kind overrides the extension', async () => {
    const copy = join(dir, 'mislabelled.docx');
    writeFileSync(copy, readFileSync(VALID_XLSX));
    expect((await validator.validateFile(copy)).ok).toBe(false);
    const res = await validator.validateFile(copy, { kind: 'xlsx' });
    expect(res.ok).toBe(true);
    expect(res.kind).toBe('xlsx');
  });
  it('an unsupported extension returns an Exception result', async () => {
    const txt = join(dir, 'notes3.txt');
    writeFileSync(txt, 'hello');
    const res = await validator.validateFile(txt);
    expect(res.ok).toBe(false);
    expect(res.errors[0].description).toBe('Unsupported extension: .txt');
  });
  it('a missing file returns an Exception result', async () => {
    const res = await validator.validateFile(join(dir, 'gone.docx'));
    expect(res.ok).toBe(false);
    expect(res.errors[0].errorType).toBe('Exception');
  });
  it('validate() takes a kind for buffers', async () => {
    const res = await validator.validate(readFileSync(VALID_PPTX), { kind: 'pptx' });
    expect(res.ok).toBe(true);
  });
  it('rejects an unknown kind', async () => {
    await expect(validator.validate(readFileSync(VALID_DOCX), { kind: 'odt' as DocumentKind })).rejects.toThrow(/Invalid kind/);
  });
});

describe('macro-enabled, template, slide show and add-in extensions', () => {
  const validator = createValidator();
  afterAll(() => validator.close());

  const bases: Record<DocumentKind, string> = { docx: VALID_DOCX, xlsx: VALID_XLSX, pptx: VALID_PPTX };
  it.each(VARIANTS)('$ext validates by file path, CLI and worker', async ({ ext, kind, type }) => {
    const file = makeVariant(bases[kind], kind, ext, type);
    expect((await validateFile(file)).ok).toBe(true);
    expect((await runCli([file])).code).toBe(0);
    const res = await validator.validateFile(file);
    expect(res.ok).toBe(true);
    expect(res.kind).toBe(kind);
  });
});
