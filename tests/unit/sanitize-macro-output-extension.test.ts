import { describe, it, expect, vi, afterEach, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

/**
 * A package that still holds macro parts has to be saved as .docm or .dotm: Word refuses a .docx that has them.
 * `sanitize --allow-macros` keeps the parts, so it refuses an output with another extension and writes nothing.
 */
let dir: string;
let seen: string[];

function macroDocx(): Buffer {
  const base = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  base.addFile('word/vbaProject.bin', Buffer.from('fake-vba-bytes', 'utf8'));
  const relsName = 'word/_rels/document.xml.rels';
  const entry = base.getEntry(relsName);
  const rels = entry
    ? entry.getData().toString('utf8')
    : '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const patched = rels.replace(
    '</Relationships>',
    '<Relationship Id="rMacro" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
  );
  if (entry) base.updateFile(relsName, Buffer.from(patched, 'utf8'));
  else base.addFile(relsName, Buffer.from(patched, 'utf8'));
  // The macro part has a content type, as in a package that Word wrote.
  const types = base.readAsText('[Content_Types].xml');
  base.updateFile(
    '[Content_Types].xml',
    Buffer.from(types.replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default '), 'utf8'),
  );
  return base.toBuffer();
}

let input: string;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-san-macro-'));
  input = path.join(dir, 'in.docm');
  fs.writeFileSync(input, macroDocx());
  seen = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('sanitize --allow-macros and the output extension', () => {
  it('refuses a .docx output, names the fix, and writes nothing', async () => {
    const out = path.join(dir, 'out.docx');
    expect(await run(['sanitize', input, out, '--allow-macros'])).toBe(2);
    const text = seen.join('\n');
    expect(text).toContain('.docm');
    expect(text).toContain('--strip-macros');
    expect(fs.existsSync(out)).toBe(false);
  });

  it('refuses a .docx output on the redaction path too', async () => {
    const out = path.join(dir, 'out.docx');
    expect(await run(['sanitize', input, out, '--allow-macros', '--redact-pii'])).toBe(2);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('refuses an output with no extension', async () => {
    const out = path.join(dir, 'out');
    expect(await run(['sanitize', input, out, '--allow-macros'])).toBe(2);
    expect(fs.existsSync(out)).toBe(false);
  });

  it('accepts .docm', async () => {
    const out = path.join(dir, 'out.docm');
    expect(await run(['sanitize', input, out, '--allow-macros'])).toBe(0);
    expect(new AdmZip(out).getEntry('word/vbaProject.bin')).not.toBeNull();
  });

  it('accepts .dotm and a capital extension', async () => {
    expect(await run(['sanitize', input, path.join(dir, 'a.dotm'), '--allow-macros'])).toBe(0);
    expect(await run(['sanitize', input, path.join(dir, 'b.DOCM'), '--allow-macros'])).toBe(0);
  });

  it('accepts a .docx output when --strip-macros removes the macro parts', async () => {
    const out = path.join(dir, 'out.docx');
    expect(await run(['sanitize', input, out, '--strip-macros'])).toBe(0);
    expect(new AdmZip(out).getEntry('word/vbaProject.bin')).toBeNull();
  });

  it('accepts a .docx output for a document that has no macro parts', async () => {
    const clean = path.join(dir, 'clean.docx');
    fs.writeFileSync(clean, DocxParser.createNewDocument('blank').rezip());
    const out = path.join(dir, 'out.docx');
    expect(await run(['sanitize', clean, out, '--allow-macros'])).toBe(0);
  });

  it('accepts a .docx output when only a relationships file points at a macro part and no macro part is stored', async () => {
    const zip = new AdmZip(macroDocx());
    zip.deleteFile('word/vbaProject.bin');
    // The relationship stays but points outside the package, so no part is missing.
    const relsName = 'word/_rels/document.xml.rels';
    zip.updateFile(
      relsName,
      Buffer.from(
        zip.readAsText(relsName).replace('Target="vbaProject.bin"', 'Target="https://example.com/vbaProject.bin" TargetMode="External"'),
        'utf8',
      ),
    );
    const dangling = path.join(dir, 'dangling.docx');
    fs.writeFileSync(dangling, zip.toBuffer());
    const out = path.join(dir, 'out.docx');
    expect(await run(['sanitize', dangling, out, '--allow-macros'])).toBe(0);
  });
});
