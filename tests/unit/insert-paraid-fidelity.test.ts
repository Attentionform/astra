import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import crypto from 'crypto';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  isValidParaId,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

function sha256Hex(input: string | Buffer): string {
  return crypto.createHash('sha256').update(input).digest('hex');
}

async function freshDoc(): Promise<{ parser: DocxParser; ast: any; anchor: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id as string;
  return { parser, ast, anchor };
}

function findNode(ast: any, id: string): any {
  const node = ast.sections
    .flatMap((s: any) => s.nodes)
    .find((n: any) => n.id === id);
  expect(node).toBeDefined();
  return node;
}

describe('inserted paragraph identity stamps', () => {
  it('inserted paragraph carries a valid w14:paraId immediately', async () => {
    const { ast, anchor } = await freshDoc();
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'stamped line', id: 'para-stamp-a',
    } as any, ast, undefined as any);
    const node = findNode(ast, 'para-stamp-a');
    const paraId = (node.xml as Element).getAttribute('w14:paraId');
    expect(paraId).toMatch(/^[0-9A-Fa-f]{8}$/);
    expect(isValidParaId(paraId)).toBe(true);
  });

  it('consecutive inserts mint distinct paraIds', async () => {
    const { ast, anchor } = await freshDoc();
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'first of two', id: 'para-stamp-b1',
    } as any, ast, undefined as any);
    const first = findNode(ast, 'para-stamp-b1');
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor: 'para-stamp-b1', position: 'after',
      text: 'second of two', id: 'para-stamp-b2',
    } as any, ast, undefined as any);
    const second = findNode(ast, 'para-stamp-b2');
    const id1 = (first.xml as Element).getAttribute('w14:paraId');
    const id2 = (second.xml as Element).getAttribute('w14:paraId');
    expect(isValidParaId(id1)).toBe(true);
    expect(isValidParaId(id2)).toBe(true);
    expect(id1).not.toBe(id2);
    expect(id1!.toUpperCase()).not.toBe(id2!.toUpperCase());
  });

  it('inserted table stamps every cell paragraph with distinct valid paraIds', async () => {
    const { ast, anchor } = await freshDoc();
    await executeInsert({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['H1', 'H2'], rows: [['a', 'b'], ['c', 'd']],
      id: 'tbl-stamp-ids',
    } as any, ast, undefined as any);
    const tblNode = findNode(ast, 'tbl-stamp-ids');
    const paras = (tblNode.xml as Element).getElementsByTagName('w:p');
    expect(paras.length).toBe(6);
    const seen = new Set<string>();
    for (let i = 0; i < paras.length; i++) {
      const pid = (paras[i] as unknown as Element).getAttribute('w14:paraId');
      expect(isValidParaId(pid), `cell paragraph ${i} carries a valid paraId`).toBe(true);
      seen.add(pid!.toUpperCase());
    }
    expect(seen.size).toBe(6);
  });

  it('caller-supplied valid paraId on xml insert is preserved', async () => {
    const { ast, anchor } = await freshDoc();
    const xml =
      `<w:p xmlns:w14="${W14}" w14:paraId="1A2B3C4D">` +
      `<w:r><w:t>kept identity</w:t></w:r></w:p>`;
    await executeInsert({
      action: 'insert', type: 'xml', anchor, position: 'after',
      xml, id: 'xml-stamp-kept',
    } as any, ast, undefined as any);
    const node = findNode(ast, 'xml-stamp-kept');
    expect((node.xml as Element).getAttribute('w14:paraId')).toBe('1A2B3C4D');
  });

  it('saved commit bytes survive a package snapshot restore byte-identically', async () => {
    const { parser, ast, anchor } = await freshDoc();
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'first committed line', id: 'commit-one',
    } as any, ast, parser);

    const commitBytes1 = parser.readDocumentXml();
    const commitHash1 = sha256Hex(commitBytes1);
    const commitSnap1 = parser.capturePackageSnapshot();

    const rezipBuf1 = parser.rezip(ast);
    const rezippedDoc1 = new AdmZip(rezipBuf1).readAsText('word/document.xml');
    expect(rezippedDoc1).toBe(commitBytes1);

    await executeInsert({
      action: 'insert', type: 'paragraph', anchor: 'commit-one', position: 'after',
      text: 'second committed line', id: 'commit-two',
    } as any, ast, parser);
    const commitBytes2 = parser.readDocumentXml();
    expect(sha256Hex(commitBytes2)).not.toBe(commitHash1);
    const commitSnap2 = parser.capturePackageSnapshot();

    parser.restorePackageSnapshot(commitSnap1);
    const restored1 = parser.readDocumentXml();
    expect(sha256Hex(restored1)).toBe(commitHash1);
    expect(restored1).toBe(commitBytes1);

    parser.restorePackageSnapshot(commitSnap2);
    expect(parser.readDocumentXml()).toBe(commitBytes2);
  });

  it('minted paraIds are deterministic across independent document instances executing identical inserts', async () => {
    const doc1 = await freshDoc();
    const doc2 = await freshDoc();

    await executeInsert({
      action: 'insert', type: 'paragraph', anchor: doc1.anchor, position: 'after',
      text: 'identical insert', id: 'deterministic-op-1',
    } as any, doc1.ast, doc1.parser);

    await executeInsert({
      action: 'insert', type: 'paragraph', anchor: doc2.anchor, position: 'after',
      text: 'identical insert', id: 'deterministic-op-1',
    } as any, doc2.ast, doc2.parser);

    const node1 = findNode(doc1.ast, 'deterministic-op-1');
    const node2 = findNode(doc2.ast, 'deterministic-op-1');

    const paraId1 = (node1.xml as Element).getAttribute('w14:paraId');
    const paraId2 = (node2.xml as Element).getAttribute('w14:paraId');

    expect(paraId1).toBeDefined();
    expect(paraId1).toBe(paraId2);

    const xml1 = doc1.parser.readDocumentXml();
    const xml2 = doc2.parser.readDocumentXml();
    expect(xml1).toBe(xml2);
  });
});
