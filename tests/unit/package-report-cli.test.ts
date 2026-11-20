/** DN3 9.2: the CLI shows the signed and the rewritten-prefix notes, in human output (stderr) and in the JSON envelope. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import AdmZip from 'adm-zip';
import { DocxParser, SIGNED_WARNING, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { EXIT_OK, run as runCli } from '../../packages/cli/src/cli';
import { seededBlank } from '../helpers/seeded-blank';

const renameW = (xml: string): string =>
  xml.replace(/(<\/?)w:/g, '$1x:').replace(/(\s)w:([A-Za-z])/g, '$1x:$2').replace(/xmlns:w=/g, 'xmlns:x=');
function build(o: { prefixed?: boolean; signed?: boolean }): Buffer {
  const zip = new AdmZip(seededBlank());
  if (o.prefixed) zip.updateFile('word/document.xml', Buffer.from(renameW(zip.readAsText('word/document.xml')), 'utf8'));
  if (o.signed) {
    zip.addFile('_xmlsignatures/origin.sigs', Buffer.alloc(0));
    zip.addFile('_xmlsignatures/sig1.xml', Buffer.from('<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>', 'utf8'));
    zip.updateFile('[Content_Types].xml', Buffer.from(zip.readAsText('[Content_Types].xml')
      .replace('<Default Extension="xml"', '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/><Default Extension="xml"'), 'utf8'));
  }
  return zip.toBuffer();
}

let dir: string;
let stdout: string[];
let stderr: string[];
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pkgcli-'));
  stdout = [];
  stderr = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => void stdout.push(a.join(' ')));
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => void stderr.push(a.join(' ')));
});
afterEach(() => { vi.restoreAllMocks(); fs.rmSync(dir, { recursive: true, force: true }); });

const write = (name: string, buf: Buffer): string => { const f = path.join(dir, name); fs.writeFileSync(f, buf); return f; };
const jsonOf = (): Record<string, unknown> => JSON.parse(stdout.filter((l) => l.trim().startsWith('{')).pop() as string);

describe('inspect', () => {
  it('a canonical, unsigned package prints no package warning and the JSON has no packageNotes', async () => {
    const f = write('plain.docx', build({}));
    expect(await runCli(['inspect', f])).toBe(EXIT_OK);
    expect(stderr.join('\n')).not.toMatch(/digitally signed|namespace prefixes/);
    stdout = [];
    expect(await runCli(['inspect', f, '--json'])).toBe(EXIT_OK);
    expect('packageNotes' in jsonOf()).toBe(false);
  });

  it('a signed, canonical package prints the signature warning (no prefix needed rewriting)', async () => {
    const f = write('signed.docx', build({ signed: true }));
    expect(await runCli(['inspect', f])).toBe(EXIT_OK);
    expect(stderr.join('\n')).toContain(SIGNED_WARNING);
    stdout = [];
    expect(await runCli(['inspect', f, '--json'])).toBe(EXIT_OK);
    const notes = jsonOf().packageNotes as { signed: boolean; warnings: string[] };
    expect(notes.signed).toBe(true);
    expect(notes.warnings).toEqual([SIGNED_WARNING]);
  });

  it('a prefixed package prints the rewrite note', async () => {
    const f = write('prefixed.docx', build({ prefixed: true }));
    expect(await runCli(['inspect', f])).toBe(EXIT_OK);
    expect(stderr.join('\n')).toMatch(/word\/document\.xml: x to w/);
  });
});

describe('batch', () => {
  it('editing a signed package saves, and the output carries the signature warning', async () => {
    const f = write('signed-edit.docx', build({ signed: true }));
    const id = (await new DocxParser(fs.readFileSync(f)).getAST()).sections[0].nodes[0].id;
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'style', target: id, color: 'FF0000' }] }));
    expect(await runCli(['batch', f, ops, '--no-history', '--json'])).toBe(EXIT_OK);
    expect(stderr.join('\n')).toContain(SIGNED_WARNING);
    const payload = jsonOf() as { ok: boolean; warnings?: string[]; packageNotes?: { signed: boolean } };
    expect(payload.ok).toBe(true);
    expect(payload.warnings).toContain(SIGNED_WARNING);
    expect(payload.packageNotes?.signed).toBe(true);
    expect(new AdmZip(fs.readFileSync(f)).readAsText('word/document.xml')).toContain('FF0000');
  });

  it('editing a prefixed package saves with the usual prefixes and says so', async () => {
    const f = write('prefixed-edit.docx', build({ prefixed: true }));
    const id = (await new DocxParser(fs.readFileSync(f)).getAST()).sections[0].nodes[0].id;
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify({ operations: [{ action: 'style', target: id, color: 'FF0000' }] }));
    expect(await runCli(['batch', f, ops, '--no-history', '--json'])).toBe(EXIT_OK);
    expect(stderr.join('\n')).toMatch(/rewritten in memory/);
    const doc = new AdmZip(fs.readFileSync(f)).readAsText('word/document.xml');
    expect(doc).toContain('<w:document');
    expect(doc).not.toMatch(/<x:/);
    expect(doc).toContain('FF0000');
  });
});
