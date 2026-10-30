import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  findMacroPartNames,
  macroChoiceError,
  macroRefusalError,
  sanitizeDocxBuffer,
  sanitizePackageEntries,
} from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';

const tmpFiles: string[] = [];

function tmpPath(prefix: string, ext = '.docx'): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}${ext}`);
  tmpFiles.push(file);
  return file;
}

function cleanDocx(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

function macroDocx(): Buffer {
  const base = new AdmZip(cleanDocx());
  base.addFile('word/vbaProject.bin', Buffer.from('fake-vba-bytes', 'utf8'));
  const relsName = 'word/_rels/document.xml.rels';
  const relsEntry = base.getEntry(relsName);
  const relsXml = relsEntry
    ? relsEntry.getData().toString('utf8')
    : '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const macroRel =
    '<Relationship Id="rMacro" ' +
    'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" ' +
    'Target="../vbaProject.bin"/>';
  const patched = relsXml.replace('</Relationships>', `${macroRel}</Relationships>`);
  if (relsEntry) base.updateFile(relsName, Buffer.from(patched, 'utf8'));
  else base.addFile(relsName, Buffer.from(patched, 'utf8'));
  return base.toBuffer();
}

function entryNames(buf: Buffer): string[] {
  const zip = new AdmZip(buf);
  return zip.getEntries().filter((e) => !e.isDirectory).map((e) => e.entryName);
}

function readText(buf: Buffer, name: string): string | null {
  const zip = new AdmZip(buf);
  const e = zip.getEntry(name);
  return e ? e.getData().toString('utf8') : null;
}

beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
});

describe('active content detection', () => {
  it('names vbaProject.bin by basename alone', () => {
    const parts = findMacroPartNames(['word/document.xml', 'word/vbaProject.bin']);
    expect(parts).toContain('word/vbaProject.bin');
  });

  it('names the word/activeX cluster by directory', () => {
    const parts = findMacroPartNames(['word/document.xml', 'word/activeX/activeX1.bin']);
    expect(parts).toContain('word/activeX/activeX1.bin');
  });

  it('counts a corroborated oleObject bin only with a rel target', () => {
    const rels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="oleObject1.bin"/>' +
      '</Relationships>';
    const read = (n: string) => (n === 'word/_rels/document.xml.rels' ? rels : null);
    expect(
      findMacroPartNames(['word/embeddings/oleObject1.bin', 'word/_rels/document.xml.rels'], read)
    ).toContain('word/embeddings/oleObject1.bin');
  });

  it('M4: refuses an unreferenced embeddings .bin by name (opaque, unscannable)', () => {
    const rels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="image1.png"/>' +
      '</Relationships>';
    const read = (n: string) => (n === 'word/_rels/document.xml.rels' ? rels : null);
    expect(
      findMacroPartNames(['word/embeddings/oleObject1.bin', 'word/_rels/document.xml.rels'], read)
    ).toContain('word/embeddings/oleObject1.bin');
  });

  it('still ignores an unreferenced oleObject bin outside embeddings (R-10 corroboration rule)', () => {
    const rels =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="image1.png"/>' +
      '</Relationships>';
    const read = (n: string) => (n === 'word/_rels/document.xml.rels' ? rels : null);
    expect(
      findMacroPartNames(['word/oleObject1.bin', 'word/_rels/document.xml.rels'], read)
    ).toEqual([]);
  });

  it('refusal names parts and offers preserve vs strip', () => {
    const err = macroRefusalError(['word/vbaProject.bin', 'word/_rels/document.xml.rels'], 'open');
    expect((err as { code?: unknown }).code).toBe('E_VALIDATION');
    expect(err.message).toContain('E_VALIDATION');
    expect(err.message).toContain('word/vbaProject.bin');
    expect(err.message).toContain('--allow-macros');
    expect(err.message).toContain('--strip-macros');
  });

  it('both choices together fail closed', () => {
    const err = macroChoiceError();
    expect((err as { code?: unknown }).code).toBe('E_VALIDATION');
    expect(err.message).toContain('--allow-macros');
  });
});

describe('sanitize gate', () => {
  it('refuses a macro package by default with zero output', () => {
    const buf = macroDocx();
    expect(() => sanitizeDocxBuffer(buf, {})).toThrowError(/E_VALIDATION/);
    expect(() => sanitizeDocxBuffer(buf, {})).toThrowError(/vbaProject\.bin/);
  });

  it('preserve keeps macro bytes verbatim', () => {
    const buf = macroDocx();
    const { buffer, report } = sanitizeDocxBuffer(buf, { allowMacros: true });
    expect(entryNames(buffer)).toContain('word/vbaProject.bin');
    expect(readText(buffer, 'word/vbaProject.bin')).toBe('fake-vba-bytes');
    expect(report.macroParts).toContain('word/vbaProject.bin');
    expect(report.strippedMacros).toEqual([]);
  });

  it('strip removes macro parts and lists removals', () => {
    const buf = macroDocx();
    const { buffer, report } = sanitizeDocxBuffer(buf, { stripMacros: true });
    expect(entryNames(buffer)).not.toContain('word/vbaProject.bin');
    expect(report.strippedMacros).toContain('word/vbaProject.bin');
    expect(report.macroParts).toContain('word/vbaProject.bin');
  });

  it('map entry point refuses without a choice', () => {
    const files = new Map<string, Buffer>([
      ['word/document.xml', Buffer.from('<w:document xmlns:w="x"/>', 'utf8')],
      ['word/vbaProject.bin', Buffer.from('x', 'utf8')],
    ]);
    expect(() => sanitizePackageEntries(files, {})).toThrowError(/E_VALIDATION/);
    expect(files.has('word/vbaProject.bin')).toBe(true);
  });
});

describe('parser open/save gate', () => {
  it('open refuses by default naming the parts', () => {
    expect(() => new DocxParser(macroDocx())).toThrowError(/E_VALIDATION/);
    expect(() => new DocxParser(macroDocx())).toThrowError(/vbaProject\.bin/);
  });

  it('open with preserve keeps the part and round-trips its bytes', () => {
    const buf = macroDocx();
    const parser = new DocxParser(buf, { allowMacros: true });
    expect(parser.getMacroParts()).toContain('word/vbaProject.bin');
    const out = new AdmZip(parser.rezip());
    expect(out.getEntry('word/vbaProject.bin')?.getData().toString('utf8')).toBe('fake-vba-bytes');
  });

  it('open with strip removes the part with accounting', () => {
    const parser = new DocxParser(macroDocx(), { stripMacros: true });
    expect(parser.getDroppedMacroEntries()).toContain('word/vbaProject.bin');
    const names = new AdmZip(parser.rezip()).getEntries().map((e) => e.entryName);
    expect(names).not.toContain('word/vbaProject.bin');
  });

  it('save refuses when the live package carries macro parts', () => {
    const parser = new DocxParser(cleanDocx());
    const zip = (parser as unknown as { zip: InstanceType<typeof AdmZip> }).zip;
    zip.addFile('word/vbaProject.bin', Buffer.from('late-add', 'utf8'));
    expect(() => parser.rezip()).toThrowError(/E_VALIDATION/);
  });

  it('clean documents behave exactly as before', () => {
    const buf = cleanDocx();
    const parser = new DocxParser(buf);
    expect(parser.getMacroParts()).toEqual([]);
    expect(parser.getDroppedMacroEntries()).toEqual([]);
  });
});

describe('cli flags', () => {
  it('sanitize refuses a macro file by default and writes nothing', async () => {
    const input = tmpPath('macro-in');
    const output = tmpPath('macro-out');
    fs.writeFileSync(input, macroDocx());
    const code = await run(['sanitize', input, output]);
    expect(code).toBe(1);
    expect(fs.existsSync(output)).toBe(false);
  });

  it('sanitize --strip-macros strips and writes a clean file', async () => {
    const input = tmpPath('macro-strip-in');
    const output = tmpPath('macro-strip-out');
    fs.writeFileSync(input, macroDocx());
    const code = await run(['sanitize', input, output, '--strip-macros']);
    expect(code).toBe(0);
    expect(entryNames(fs.readFileSync(output))).not.toContain('word/vbaProject.bin');
  });

  it('sanitize with both flags is a usage error', async () => {
    const input = tmpPath('macro-both-in');
    const output = tmpPath('macro-both-out');
    fs.writeFileSync(input, macroDocx());
    const code = await run(['sanitize', input, output, '--allow-macros', '--strip-macros']);
    expect(code).toBe(2);
    expect(fs.existsSync(output)).toBe(false);
  });

  it('validate refuses by default and passes with strip', async () => {
    const file = tmpPath('macro-validate');
    fs.writeFileSync(file, macroDocx());
    expect(await run(['validate', file])).toBe(1);
    expect(await run(['validate', file, '--strip-macros'])).toBe(0);
  });

  it('inspect refuses by default and reads with preserve', async () => {
    const file = tmpPath('macro-inspect');
    fs.writeFileSync(file, macroDocx());
    expect(await run(['inspect', file])).toBe(1);
    expect(await run(['inspect', file, '--allow-macros'])).toBe(0);
  });
});

describe('session store gate', () => {
  it('openBuffer refuses by default and opens with an explicit choice', async () => {
    const store = createMcpSessionStore();
    await expect(store.openBuffer(macroDocx(), 'macro.docx')).rejects.toMatchObject({
      code: 'E_VALIDATION',
    });
    const kept = await store.openBuffer(macroDocx(), 'macro.docx', undefined, { allowMacros: true });
    expect(kept.parser.getMacroParts()).toContain('word/vbaProject.bin');
    const stripped = await store.openBuffer(macroDocx(), 'macro.docx', undefined, { stripMacros: true });
    expect(stripped.parser.getDroppedMacroEntries()).toContain('word/vbaProject.bin');
  });
});
