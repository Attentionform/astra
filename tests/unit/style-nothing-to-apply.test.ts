import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser } from '../../packages/engine/src';
import { executeStyle, type StyleOperation } from '../../packages/engine/src/operations/style';

const GALLERY = path.resolve(__dirname, '../../test-gallery/T2_edit_target.docx');

async function setup() {
  const parser = new DocxParser(fs.readFileSync(GALLERY));
  const ast = await parser.getAST();
  const para = ast.sections[0].nodes.find((n) => n.type === 'paragraph')!;
  const word = String(para.metadata.text ?? '').split(/\s+/)[0];
  return { parser, ast, id: para.id, word };
}

async function attempt(op: (id: string, word: string) => Record<string, unknown>) {
  const { parser, ast, id, word } = await setup();
  const before = parser.readDocumentXml();
  let error: { code?: string; message?: string; remedy?: string } | undefined;
  try {
    executeStyle({ action: 'style', target: id, ...op(id, word) } as unknown as StyleOperation, ast);
  } catch (err) {
    error = err as { code?: string; message?: string; remedy?: string };
  }
  return { error, changed: parser.readDocumentXml() !== before };
}

describe('a style op must change something or say why not', () => {
  it('refuses a match-only op (it selects text but sets no property)', async () => {
    const r = await attempt((_id, word) => ({ match: word }));
    expect(r.error?.code).toBe('E_VALIDATION');
    expect(r.error?.message).toContain('no style property');
    expect(r.changed).toBe(false);
  });

  it('refuses an op with no properties at all, flat or nested', async () => {
    for (const op of [{}, { properties: {} }, { params: {} }, { matchAll: true }]) {
      const r = await attempt(() => op);
      expect(r.error?.code, JSON.stringify(op)).toBe('E_VALIDATION');
      expect(r.changed).toBe(false);
    }
  });

  it('still applies a real property, with or without match', async () => {
    const plain = await attempt(() => ({ bold: true }));
    expect(plain.error).toBeUndefined();
    expect(plain.changed).toBe(true);
    const matched = await attempt((_id, word) => ({ bold: true, match: word }));
    expect(matched.error).toBeUndefined();
    expect(matched.changed).toBe(true);
  });
});

describe('a border config only reads style, sz and color', () => {
  it('refuses an unknown key inside a border config, wherever the border sits', async () => {
    const flat = await attempt(() => ({ borderBox: { style: 'single', sz: 8, color: '000000', wat: 1 } }));
    expect(flat.error?.code).toBe('E_VALIDATION');
    expect(flat.error?.message).toContain('borderBox.wat');
    expect(flat.changed).toBe(false);
    const nested = await attempt(() => ({ properties: { borderTop: { style: 'single', wat: 1 } } }));
    expect(nested.error?.message).toContain('properties.borderTop.wat');
    const params = await attempt(() => ({ bold: true, params: { borderLeft: { colour: 'FF0000' } } }));
    expect(params.error?.message).toContain('params.borderLeft.colour');
    expect(params.changed).toBe(false);
  });

  it('still applies a valid border config', async () => {
    const ok = await attempt(() => ({ borderBox: { style: 'single', sz: 8, color: '336699' } }));
    expect(ok.error).toBeUndefined();
    expect(ok.changed).toBe(true);
  });
});
