import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  RelationshipsManipulator,
  ToonGenerator,
} from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

async function setup(template: string = 'blank') {
  const parser = DocxParser.createNewDocument(template);
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  return { parser, ast, engine };
}

function documentXmlOf(parser: DocxParser, buffer?: Buffer): string {
  const zip = new AdmZip(buffer ?? parser.rezip());
  return zip.getEntry('word/document.xml')!.getData().toString('utf8');
}

describe('Engine transactional guarantees', () => {
  it('(a) failed move batch leaves live DOM byte-identical and a later insert lands in rezip output', async () => {
    const { parser, ast, engine } = await setup('blank');
    const before = parser.readDocumentXml();
    const P1 = ast.sections[0].nodes[0].id;
    const P2 = ast.sections[0].nodes[1].id;

    // Op 1 succeeds (detaches P2), op 2 detaches P1 then throws on the
    // unknown anchor — the live DOM would be gutted without a real rollback.
    const res = await engine.executeBatch({
      operations: [
        { action: 'move', target: P2, anchor: P1, position: 'before' },
        { action: 'move', target: P1, anchor: 'nope_missing', position: 'after' }
      ],
      sequential: true
    });
    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);

    // Byte-identical restoration of the live document
    expect(parser.readDocumentXml()).toBe(before);

    // A subsequent successful insert lands in the actual rezip output
    const retry = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: P1, position: 'before', text: 'TRANSACTION PROOF' }
      ]
    });
    expect(retry.success).toBe(true);
    expect(documentXmlOf(parser)).toContain('TRANSACTION PROOF');
  });

  it('rejects duplicate custom operation ids inside one batch with a structured error', async () => {
    const { ast, engine } = await setup('blank');
    const P1 = ast.sections[0].nodes[0].id;
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: P1, position: 'before', text: 'one', id: 'dup-1' },
        { action: 'insert', type: 'paragraph', anchor: P1, position: 'before', text: 'two', id: 'dup-1' }
      ]
    });
    expect(res.success).toBe(false);
    expect(res.appliedCount).toBe(0);
    expect(res.errors[0].error).toContain("Duplicate operation id 'dup-1'");
    expect(res.errors[0].operationIndex).toBe(1);
  });

  it('(b) relationship allocation on a fresh blank template yields rId >= rId2 and leaves the styles rel untouched', async () => {
    const { parser, ast, engine } = await setup('blank');

    // getAST() now populates real relationships from word/_rels/document.xml.rels
    expect(ast.relationships['rId1']).toMatchObject({ target: 'styles.xml' });

    // Header/footer-style allocation through the engine must not recycle rId1
    const res = await engine.executeBatch({
      operations: [{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'Engine Test Header' }]
    });
    expect(res.success).toBe(true);

    const allocatedRIds = Object.keys(ast.relationships).filter(r => r !== 'rId1');
    expect(allocatedRIds.length).toBeGreaterThan(0);
    for (const rId of allocatedRIds) {
      expect(parseInt(rId.replace('rId', ''), 10)).toBeGreaterThanOrEqual(2);
    }

    // styles.xml relationship still intact in the serialized package
    const relsXml = new AdmZip(parser.rezip()).getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
    expect(relsXml).toContain('Id="rId1"');
    expect(relsXml).toContain('Target="styles.xml"');

    // Direct manipulator allocation agrees
    expect(new RelationshipsManipulator(ast).generateSafeRId()).toMatch(/^rId([3-9]|\d{2,})$/);
  });

  it('(c) image insert strips the data URL prefix and writes the media part plus Relationship unconditionally', async () => {
    const { parser, ast, engine } = await setup('blank');
    const pngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

    const res = await engine.executeBatch({
      operations: [
        { action: 'image', operation: 'insert', anchor: ast.sections[0].nodes[1].id, fileData: `data:image/png;base64,${pngBase64}` }
      ]
    });
    expect(res.success).toBe(true);

    const mediaNames = Object.keys(ast.mediaFiles);
    expect(mediaNames.length).toBe(1);
    expect(mediaNames[0].endsWith('.png')).toBe(true);
    expect(ast.mediaFiles[mediaNames[0]].equals(Buffer.from(pngBase64, 'base64'))).toBe(true);

    // Media part really exists in the package under that name
    const zipAdapter = parser.getZipAdapter();
    expect(zipAdapter.hasFile(`word/media/${mediaNames[0]}`)).toBe(true);

    // Relationship written under a safe (non-rId1) id
    const relsXml = zipAdapter.readText('word/_rels/document.xml.rels');
    expect(relsXml).toContain(`Target="media/${mediaNames[0]}"`);
    const embedRId = Object.keys(ast.relationships).find(r => ast.relationships[r].target === `media/${mediaNames[0]}`);
    expect(embedRId).toBeDefined();
    expect(parseInt(embedRId!.replace('rId', ''), 10)).toBeGreaterThanOrEqual(2);

    // Drawing references the safe rId in the saved package
    expect(documentXmlOf(parser)).toContain(`r:embed="${embedRId}"`);
  });

  it('(d) insert with a typo/stale anchor throws instead of silently appending at the tail', async () => {
    const { parser, engine } = await setup('blank');
    const before = parser.readDocumentXml();

    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'p_999_typo', position: 'before', text: 'should not appear' }
      ]
    });
    expect(res.success).toBe(false);
    const combined = `${res.errors[0].error} ${res.errors[0].hint ?? ''}`;
    expect(combined).toContain('p_999_typo');
    expect(combined).toContain('inspect_document');

    // Nothing landed in the DOM
    expect(parser.readDocumentXml()).toBe(before);
    expect(parser.readDocumentXml()).not.toContain('should not appear');
  });

  it('(e) table updateCell bumps version so a second generateToon reflects the new text', async () => {
    const { ast, engine } = await setup('financial');
    const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
    expect(tableNode).toBeDefined();

    const generator = new ToonGenerator();
    const before = generator.generateToon(ast);
    const beforeRow = before.rows.find(r => r.id === tableNode!.id)!;
    expect(beforeRow.text).toContain('Net Revenue');

    const res = await engine.executeBatch({
      operations: [
        {
          action: 'table',
          target: tableNode!.id,
          params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'UPDATED REVENUE' }
        }
      ]
    });
    expect(res.success).toBe(true);

    expect(tableNode!.version).toBeGreaterThan(1);

    // Second generation (with cache enabled) must reflect the mutated cell
    const after = generator.generateToon(ast);
    const afterRow = after.rows.find(r => r.id === tableNode!.id)!;
    expect(afterRow.text).toContain('UPDATED REVENUE');
    expect(afterRow.version).toBeGreaterThan(beforeRow.version);
  });

  it('(f) created w:t carries xml:space=preserve for text with leading/trailing whitespace', async () => {
    const { parser, ast, engine } = await setup('blank');
    const res = await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'default', text: '  padded body text  ' }
      ]
    });
    expect(res.success).toBe(true);

    const newNode = ast.sections[0].nodes.find(n => n.metadata.text === '  padded body text  ');
    expect(newNode).toBeDefined();

    const serialized = parser.readNode(newNode!.id);
    expect(serialized).toContain('<w:t xml:space="preserve">  padded body text  </w:t>');

    // And it survives the round-trip through the saved package
    expect(documentXmlOf(parser)).toContain('<w:t xml:space="preserve">  padded body text  </w:t>');
  });

  it('pageSetup mutates <w:sectPr> for real with schema-correct child order', async () => {
    const { parser, engine } = await setup('blank');
    const res = await engine.executeBatch({
      operations: [
        {
          action: 'pageSetup',
          params: { paperSize: 'A4', orientation: 'landscape', marginTop: 720, columns: 2 }
        }
      ]
    });
    expect(res.success).toBe(true);

    const xml = documentXmlOf(parser);
    // A4 landscape: width 16838 x height 11906 (swapped)
    expect(xml).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>');
    expect(xml).toContain('w:top="720"');
    expect(xml).toContain('<w:cols w:num="2"/>');

    // Schema order inside sectPr: pgSz before pgMar before cols
    const sectPrMatch = xml.match(/<w:sectPr[^>]*>([\s\S]*?)<\/w:sectPr>/);
    expect(sectPrMatch).toBeTruthy();
    const sectPr = sectPrMatch![1];
    const iPgSz = sectPr.indexOf('<w:pgSz');
    const iPgMar = sectPr.indexOf('<w:pgMar');
    const iCols = sectPr.indexOf('<w:cols');
    expect(iPgSz).toBeGreaterThanOrEqual(0);
    expect(iPgMar).toBeGreaterThan(iPgSz);
    expect(iCols).toBeGreaterThan(iPgMar);
  });

  it('unsupported sub-actions throw instead of logging success', async () => {
    const { ast, engine } = await setup('blank');
    const P1 = ast.sections[0].nodes[0].id;
    // NOTE (D2 OX-02): `equation` was implemented (minimal OMML) and left
    // this failure map — it now succeeds and is pinned in ops-d2.test.ts.
    const cases: Record<string, object[]> = {
      'section.unsupported_op is not supported yet': [
        { action: 'section', operation: 'unsupported_op' as any, target: P1 }
      ],
      'field.unsupported_op is not supported yet': [
        { action: 'field', operation: 'unsupported_op' as any, anchor: P1 }
      ],
      'equation.bogus-op is not supported yet': [
        { action: 'equation', operation: 'bogus-op' as any, anchor: P1, latex: 'E=mc^2' }
      ],
      'namedStyle.create is not supported yet': [
        { action: 'namedStyle', operation: 'create', styleName: 'Callout' }
      ],
      'image.unsupported is not supported yet': [
        { action: 'image', operation: 'unsupported' as any, target: 'img_x', width: 100 }
      ],
    };
    for (const [expected, operations] of Object.entries(cases)) {
      const res = await engine.executeBatch({ operations: operations as any[] });
      expect(res.success, expected).toBe(false);
      expect(res.rollbackTriggered, expected).toBe(true);
      expect(res.errors.some(e => e.error.includes(expected)), expected).toBe(true);
    }

    // Unsupported table op test
    const financial = await setup('financial');
    const tableRes = await financial.engine.executeBatch({
      operations: [
        { action: 'table', target: 't_001', params: { op: 'unsupported_op' } }
      ]
    });
    expect(tableRes.success).toBe(false);
    expect(tableRes.errors.some(e => e.error.includes('table.unsupported_op is not supported yet'))).toBe(true);
    // Direct call surfaces the same honesty rule
    expect(() =>
      engine.execute({ action: 'list', operation: 'create', target: 'missing_node' } as any)
    ).rejects.toThrow(/missing_node/);
  });

  it('replace no longer retargets unknown targets to the first node', async () => {
    const { parser, engine } = await setup('blank');
    const before = parser.readDocumentXml();
    const res = await engine.executeBatch({
      operations: [{ action: 'replace', target: 'stale_id', text: 'MISFIRE' }]
    });
    expect(res.success).toBe(false);
    expect(res.errors[0].error).toContain('stale_id');
    expect(parser.readDocumentXml()).toBe(before);
  });
});
