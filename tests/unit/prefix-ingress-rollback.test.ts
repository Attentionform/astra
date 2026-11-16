/**
 * DN3 Phase 0 spike: a package whose WordprocessingML uses another prefix goes through the ingress at open,
 * edits, undo, redo and a refused save; the original bytes stay exact and recoverable.
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocumentManager, DocxParser, createBlankDocxBuffer } from '@attentionform/astra-engine';
import { getOriginalInfo, originalSha256 } from '../../packages/engine/src/original';
import { verifyEquivalent } from '../../packages/engine/src/ooxml/prefix-canonicaliser';
import { createHash } from 'node:crypto';

const renameW = (xml: string): string =>
  xml.replace(/(<\/?)w:/g, '$1x:').replace(/(\s)w:([A-Za-z])/g, '$1x:$2').replace(/xmlns:w=/g, 'xmlns:x=');

function prefixedDocx(): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer());
  for (const e of zip.getEntries()) {
    if (e.isDirectory || !/^word\/.*\.xml$/.test(e.entryName)) continue;
    zip.updateFile(e.entryName, Buffer.from(renameW(e.getData().toString('utf8')), 'utf8'));
  }
  return zip.toBuffer();
}
const partText = (buf: Buffer, name: string): string => new AdmZip(buf).getEntry(name)!.getData().toString('utf8');
const sha = (b: Buffer): string => createHash('sha256').update(b).digest('hex');

let dir: string;
beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dn3-')); });
afterEach(() => { fs.rmSync(dir, { recursive: true, force: true }); });

describe('prefixed package through the ingress', () => {
  it('fixture really is prefixed', () => {
    const doc = partText(prefixedDocx(), 'word/document.xml');
    expect(doc).toContain('<x:document');
    expect(doc).not.toMatch(/<w:/);
  });

  it('DocxParser opens it, reports the rewrite, and rezip writes the usual prefix', () => {
    const parser = new DocxParser(prefixedDocx());
    expect(parser.canonicalisedParts.map((p) => p.part)).toContain('word/document.xml');
    expect(parser.getStructuralSummary().nodes.length).toBeGreaterThan(0);
    const out = parser.rezip();
    expect(partText(out, 'word/document.xml')).toContain('<w:document');
  });

  it('open, edit, undo, redo, save: original bytes retained and exact; every state canonical', async () => {
    const file = path.join(dir, 'prefixed.docx');
    const original = prefixedDocx();
    fs.writeFileSync(file, original);
    const dm = new DocumentManager();
    await dm.openDocument(file);
    const info = await getOriginalInfo(file);
    if (!info.available) throw new Error('original not retained');
    expect(info.sha256).toBe(originalSha256(original));
    expect(info.sha256).toBe(sha(original));

    const id = dm.getParser().getStructuralSummary().nodes[0].id;
    await dm.updateNode(id, '<w:p><w:r><w:t>edited</w:t></w:r></w:p>', 'edit');
    expect(partText(fs.readFileSync(file), 'word/document.xml')).toContain('edited');
    expect(partText(fs.readFileSync(file), 'word/document.xml')).toContain('<w:document');
    expect(dm.canUndo()).toBe(true);

    await dm.undoLastAction();
    expect(partText(fs.readFileSync(file), 'word/document.xml')).not.toContain('edited');
    expect(dm.canRedo()).toBe(true);
    await dm.redoLastAction();
    expect(partText(fs.readFileSync(file), 'word/document.xml')).toContain('edited');

    await dm.saveDocument();
    const after = await getOriginalInfo(file);
    if (!after.available) throw new Error('original lost');
    expect(after.sha256).toBe(sha(original)); // original untouched by every save
  });

  it('a snapshot that carries prefixed parts is canonicalised on restore', () => {
    const parser = new DocxParser(createBlankDocxBuffer());
    const snap = new Map<string, Buffer>();
    for (const e of new AdmZip(prefixedDocx()).getEntries()) if (!e.isDirectory) snap.set(e.entryName, e.getData());
    parser.restorePackageSnapshot(snap);
    expect(parser.getStructuralSummary().nodes.length).toBeGreaterThan(0);
    expect(partText(parser.rezip(), 'word/document.xml')).toContain('<w:document');
  });

  it('reloadDocumentXml accepts prefixed document text', () => {
    const parser = new DocxParser(createBlankDocxBuffer());
    const prefixed = renameW(parser.readDocumentXml());
    parser.reloadDocumentXml(prefixed);
    expect(parser.readDocumentXml()).toContain('<w:body');
  });

  it('a conflicting package is refused at open and nothing is written', async () => {
    const zip = new AdmZip(createBlankDocxBuffer());
    const xml = zip.getEntry('word/document.xml')!.getData().toString('utf8')
      .replace('<w:document ', '<w:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ')
      .replace('<w:body>', '<w:body><q:thing xmlns:q="urn:q" xmlns:w="urn:other"><x:p/></q:thing>');
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const file = path.join(dir, 'conflict.docx');
    fs.writeFileSync(file, zip.toBuffer());
    const before = sha(fs.readFileSync(file));
    const dm = new DocumentManager();
    await expect(dm.openDocument(file)).rejects.toThrow(/cannot be canonicalised safely/);
    expect(sha(fs.readFileSync(file))).toBe(before);
  });

  it('canonical result is equivalent to the prefixed original (document part)', () => {
    const original = prefixedDocx();
    const parser = new DocxParser(original);
    const check = verifyEquivalent(partText(original, 'word/document.xml'), partText(parser.rezip(), 'word/document.xml'));
    // rezip may add declarations or change attribute order but meaning of the document node set is the same
    expect(check.ok, check.problem).toBe(true);
  });

  it('refused save: restoreLastPersisted rewinds to the (prefixed) bytes read at open, canonicalised', async () => {
    const file = path.join(dir, 'refused.docx');
    fs.writeFileSync(file, prefixedDocx());
    const dm = new DocumentManager();
    await dm.openDocument(file);
    const id = dm.getParser().getStructuralSummary().nodes[0].id;
    dm.beginTransaction('t');
    dm.getParser().reloadDocumentXml(dm.getParser().readDocumentXml().replace('</w:body>', '<w:p><w:r><w:t>unsaved</w:t></w:r></w:p></w:body>'));
    expect(dm.getParser().readDocumentXml()).toContain('unsaved');
    (dm as unknown as { restoreLastPersisted(): void }).restoreLastPersisted();
    expect(dm.getParser().readDocumentXml()).not.toContain('unsaved');
    expect(dm.getParser().readDocumentXml()).toContain('<w:body');
    expect(dm.getParser().getStructuralSummary().nodes.some((n) => n.id === id)).toBe(true);
  });

  it('failed batch: rollbackTransaction restores the pre-batch state', async () => {
    const file = path.join(dir, 'batch.docx');
    fs.writeFileSync(file, prefixedDocx());
    const dm = new DocumentManager();
    await dm.openDocument(file);
    dm.beginTransaction('batch');
    const before = dm.getParser().readDocumentXml();
    dm.getParser().reloadDocumentXml(before.replace('</w:body>', '<w:p><w:r><w:t>half</w:t></w:r></w:p></w:body>'));
    dm.rollbackTransaction();
    expect(dm.getParser().readDocumentXml()).toBe(before);
  });
});
