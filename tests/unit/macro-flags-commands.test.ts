import { describe, it, expect, vi, afterEach, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

/**
 * A document with macro parts is refused unless the caller chooses. The refusal names `--allow-macros` and
 * `--strip-macros`, so every command that opens a document has to take both flags: stats, search, export,
 * normalize and the history commands. The two flags together are a usage error (exit 2).
 */
let dir: string;
let macroDoc: string;
let seen: string[];

const text = (): string => seen.join('\n');

function macroDocx(): Buffer {
  const base = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  base.addFile('word/vbaProject.bin', Buffer.from('fake-vba-bytes', 'utf8'));
  const relsName = 'word/_rels/document.xml.rels';
  const patched = base
    .readAsText(relsName)
    .replace(
      '</Relationships>',
      '<Relationship Id="rMacro" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
    );
  base.updateFile(relsName, Buffer.from(patched, 'utf8'));
  const types = base.readAsText('[Content_Types].xml');
  base.updateFile(
    '[Content_Types].xml',
    Buffer.from(types.replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default '), 'utf8'),
  );
  return base.toBuffer();
}

const hasMacroPart = (file: string): boolean => new AdmZip(file).getEntry('word/vbaProject.bin') !== null;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-macro-cmds-'));
  macroDoc = path.join(dir, 'in.docm');
  fs.writeFileSync(macroDoc, macroDocx());
  seen = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void seen.push(a.join(' ')));
});

afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('stats', () => {
  it('refuses a macro document and names both flags', async () => {
    expect(await run(['stats', macroDoc])).toBe(1);
    expect(text()).toContain('--allow-macros');
    expect(text()).toContain('--strip-macros');
  });
  it('accepts --allow-macros and --strip-macros', async () => {
    expect(await run(['stats', macroDoc, '--allow-macros'])).toBe(0);
    expect(await run(['stats', macroDoc, '--strip-macros'])).toBe(0);
    expect(text()).toContain('STATS');
  });
  it('takes the flag before the file too', async () => {
    expect(await run(['stats', '--allow-macros', macroDoc])).toBe(0);
  });
  it('refuses both flags together with exit 2', async () => {
    expect(await run(['stats', macroDoc, '--allow-macros', '--strip-macros'])).toBe(2);
    expect(text()).toContain('mutually exclusive');
  });
  it('still refuses an unknown flag and a missing file', async () => {
    expect(await run(['stats', macroDoc, '--nope'])).toBe(2);
    expect(await run(['stats', '--allow-macros'])).toBe(2);
    expect(text()).toContain('stats requires exactly one <docx> argument');
  });
});

describe('search', () => {
  it('refuses by default and accepts each flag', async () => {
    expect(await run(['search', macroDoc, 'the'])).toBe(1);
    expect(text()).toContain('--allow-macros');
    expect(await run(['search', macroDoc, 'the', '--allow-macros'])).toBe(0);
    expect(await run(['search', macroDoc, 'the', '--strip-macros'])).toBe(0);
  });
  it('refuses both flags together with exit 2', async () => {
    expect(await run(['search', macroDoc, 'the', '--allow-macros', '--strip-macros'])).toBe(2);
  });
});

describe('export', () => {
  it('refuses by default and writes nothing', async () => {
    const out = path.join(dir, 'out.html');
    expect(await run(['export', macroDoc, '--out', out])).toBe(1);
    expect(text()).toContain('--strip-macros');
    expect(fs.existsSync(out)).toBe(false);
  });
  it('exports HTML and Markdown with either flag', async () => {
    const html = path.join(dir, 'a.html');
    const md = path.join(dir, 'a.md');
    expect(await run(['export', macroDoc, '--allow-macros', '--out', html])).toBe(0);
    expect(await run(['export', macroDoc, '--strip-macros', '-f', 'md', '--out', md])).toBe(0);
    expect(fs.existsSync(html)).toBe(true);
    expect(fs.existsSync(md)).toBe(true);
  });
  it('refuses both flags together with exit 2 and writes nothing', async () => {
    const out = path.join(dir, 'b.html');
    expect(await run(['export', macroDoc, '--allow-macros', '--strip-macros', '--out', out])).toBe(2);
    expect(fs.existsSync(out)).toBe(false);
  });
  it('lists the flags in its help', async () => {
    await run(['export', '--help']);
    expect(text()).toContain('--allow-macros');
    expect(text()).toContain('--strip-macros');
  });
});

describe('normalize', () => {
  it('refuses by default with the flags in the remedy and writes nothing', async () => {
    const out = path.join(dir, 'out.docm');
    expect(await run(['normalize', macroDoc, '--out', out])).toBe(1);
    expect(text()).toContain('--allow-macros');
    expect(fs.existsSync(out)).toBe(false);
  });
  it('--dry-run takes both flags', async () => {
    expect(await run(['normalize', macroDoc, '--dry-run', '--allow-macros'])).toBe(0);
    expect(await run(['normalize', macroDoc, '--dry-run', '--strip-macros'])).toBe(0);
  });
  it('--allow-macros keeps the macro parts in a .docm output', async () => {
    const out = path.join(dir, 'out.docm');
    expect(await run(['normalize', macroDoc, '--allow-macros', '--out', out])).toBe(0);
    expect(hasMacroPart(out)).toBe(true);
  });
  it('--allow-macros refuses a .docx output and writes nothing', async () => {
    const out = path.join(dir, 'out.docx');
    expect(await run(['normalize', macroDoc, '--allow-macros', '--out', out])).toBe(2);
    expect(text()).toContain('.docm');
    expect(fs.existsSync(out)).toBe(false);
  });
  it('--strip-macros writes a .docx without the macro parts', async () => {
    const out = path.join(dir, 'out.docx');
    expect(await run(['normalize', macroDoc, '--strip-macros', '--out', out])).toBe(0);
    expect(hasMacroPart(out)).toBe(false);
    const types = new AdmZip(out).readAsText('[Content_Types].xml');
    expect(types).not.toContain('macroEnabled');
  });
  it('--strip-macros refuses to write over the .docm in place, and leaves it as it was', async () => {
    const before = fs.readFileSync(macroDoc);
    expect(await run(['normalize', macroDoc, '--strip-macros'])).toBe(2);
    expect(text()).toContain('--out <name>.docx');
    expect(fs.readFileSync(macroDoc).equals(before)).toBe(true);
  });
  it('--strip-macros refuses a .dotm output', async () => {
    const out = path.join(dir, 'out.dotm');
    expect(await run(['normalize', macroDoc, '--strip-macros', '--out', out])).toBe(2);
    expect(fs.existsSync(out)).toBe(false);
  });
  it('writes in place with --allow-macros, so the file stays a .docm with its macro parts', async () => {
    expect(await run(['normalize', macroDoc, '--allow-macros'])).toBe(0);
    expect(hasMacroPart(macroDoc)).toBe(true);
  });
  it('refuses both flags together with exit 2', async () => {
    expect(await run(['normalize', macroDoc, '--allow-macros', '--strip-macros'])).toBe(2);
  });
  it('a document without macro parts is unchanged by the new flags', async () => {
    const clean = path.join(dir, 'clean.docx');
    fs.writeFileSync(clean, DocxParser.createNewDocument('blank').rezip());
    expect(await run(['normalize', clean, '--allow-macros', '--out', path.join(dir, 'c1.docx')])).toBe(0);
    expect(await run(['normalize', clean, '--strip-macros', '--out', path.join(dir, 'c2.docx')])).toBe(0);
  });
});

describe('history commands', () => {
  it('history --init refuses by default and takes either flag', async () => {
    expect(await run(['history', macroDoc, '--init'])).toBe(1);
    expect(text()).toContain('--allow-macros');
    expect(await run(['history', macroDoc, '--init', '--allow-macros'])).toBe(0);
  });

  it('preview-revert, revert and redo take --allow-macros and keep the macro parts', async () => {
    expect(await run(['history', macroDoc, '--init', '--allow-macros'])).toBe(0);
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(
      ops,
      JSON.stringify({ operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'history step' }] }),
    );
    expect(await run(['batch', macroDoc, ops, '--allow-macros'])).toBe(0);
    expect(new AdmZip(macroDoc).readAsText('word/document.xml')).toContain('history step');

    expect(await run(['preview-revert', macroDoc])).toBe(1);
    expect(text()).toContain('--allow-macros');
    expect(await run(['preview-revert', macroDoc, '--allow-macros'])).toBe(0);

    expect(await run(['revert', macroDoc])).toBe(1);
    expect(await run(['revert', macroDoc, '--allow-macros'])).toBe(0);
    expect(hasMacroPart(macroDoc)).toBe(true);
    expect(new AdmZip(macroDoc).readAsText('word/document.xml')).not.toContain('history step');

    expect(await run(['redo', macroDoc])).toBe(1);
    expect(await run(['redo', macroDoc, '--allow-macros'])).toBe(0);
    expect(hasMacroPart(macroDoc)).toBe(true);
    expect(new AdmZip(macroDoc).readAsText('word/document.xml')).toContain('history step');
  });

  it('revert and redo refuse --strip-macros on a .docm and leave it as it was', async () => {
    expect(await run(['history', macroDoc, '--init', '--allow-macros'])).toBe(0);
    const before = fs.readFileSync(macroDoc);
    expect(await run(['revert', macroDoc, '--strip-macros'])).toBe(2);
    expect(text()).toContain('--allow-macros');
    expect(await run(['redo', macroDoc, '--strip-macros'])).toBe(2);
    expect(fs.readFileSync(macroDoc).equals(before)).toBe(true);
  });

  it('revert and redo refuse both flags together with exit 2', async () => {
    expect(await run(['revert', macroDoc, '--allow-macros', '--strip-macros'])).toBe(2);
    expect(await run(['redo', macroDoc, '--allow-macros', '--strip-macros'])).toBe(2);
    expect(await run(['preview-revert', macroDoc, '--allow-macros', '--strip-macros'])).toBe(2);
    expect(await run(['history', macroDoc, '--allow-macros', '--strip-macros'])).toBe(2);
  });
});
