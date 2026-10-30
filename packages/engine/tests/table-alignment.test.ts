import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { executeTable } from '../src/operations/table';
import { createWordElement } from '../src/ooxml/ns';
import { TBLPR_CHILD_ORDER } from '../src/ooxml/schema-order';
import { validateOperationStrict } from '../src/operation-schema';
import type { Node } from '../src/ast';

function directChildren(parent: Element, tagName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1 && (child as Element).tagName === tagName) {
      out.push(child as Element);
    }
  }
  return out;
}

async function freshDoc() {
  const parser = new DocxParser(createBlankDocxBuffer('Table Alignment Suite'));
  const ast = await parser.getAST();
  return { parser, ast };
}

function lastNodeId(ast: { sections: Array<{ nodes: Node[] }> }): string {
  const nodes = ast.sections[0].nodes;
  return nodes[nodes.length - 1].id;
}

function tableNodeOf(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const found = ast.sections[0].nodes.find((n) => n.type === 'table');
  if (!found) throw new Error('test setup: no table node found');
  return found;
}

async function insertTable(engine: OperationEngine, anchor: string): Promise<string> {
  await engine.execute({
    action: 'insert',
    type: 'table',
    position: 'after',
    anchor,
    headers: ['Header A', 'Header B'],
    rows: [
      ['Cell 1', 'Cell 2'],
      ['Cell 3', 'Cell 4'],
    ],
  } as never);
  const ast = (engine as unknown as { ast: { sections: Array<{ nodes: Node[] }> } }).ast;
  return tableNodeOf(ast).id;
}

describe('table page alignment', () => {
  it('applies left, center, and right alignment to a table', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;

    for (const alignment of ['center', 'right', 'left'] as const) {
      await engine.execute({
        action: 'table',
        target: tableId,
        operation: 'setTableAlignment',
        align: alignment,
      } as never);

      const tblPr = directChildren(tblEl, 'w:tblPr')[0];
      expect(tblPr).toBeDefined();
      const jcList = directChildren(tblPr, 'w:jc');
      expect(jcList).toHaveLength(1);
      expect(jcList[0].getAttribute('w:val')).toBe(alignment);
    }
  });

  it('accepts alignment parameter via alternate aliases and nested params', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;

    // Test top-level alignment property
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      alignment: 'center',
    } as never);

    let tblPr = directChildren(tblEl, 'w:tblPr')[0];
    let jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('center');

    // Test nested params.align
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      params: { align: 'right' },
    } as never);

    tblPr = directChildren(tblEl, 'w:tblPr')[0];
    jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('right');

    // Test nested params.alignment
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      params: { alignment: 'left' },
    } as never);

    tblPr = directChildren(tblEl, 'w:tblPr')[0];
    jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('left');
  });

  it('handles case-insensitive alignment values such as Center and RIGHT', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;

    await executeTable(
      {
        action: 'table',
        target: tableId,
        operation: 'setTableAlignment',
        align: 'Center',
      },
      ast
    );

    let tblPr = directChildren(tblEl, 'w:tblPr')[0];
    let jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('center');

    await executeTable(
      {
        action: 'table',
        target: tableId,
        operation: 'setTableAlignment',
        align: 'RIGHT',
      },
      ast
    );

    tblPr = directChildren(tblEl, 'w:tblPr')[0];
    jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('right');

    await executeTable(
      {
        action: 'table',
        target: tableId,
        operation: 'setTableAlignment',
        params: { align: 'Left' },
      },
      ast
    );

    tblPr = directChildren(tblEl, 'w:tblPr')[0];
    jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('left');
  });

  it('validates case-insensitive table alignment via validateOperationStrict', () => {
    const centerResult = validateOperationStrict({
      action: 'table',
      target: 'tbl-1',
      operation: 'setTableAlignment',
      align: 'CENTER',
    });
    expect(centerResult.ok).toBe(true);
    expect((centerResult.operation as { align?: string }).align).toBe('center');

    const rightResult = validateOperationStrict({
      action: 'table',
      target: 'tbl-1',
      operation: 'setTableAlignment',
      alignment: 'RIGHT',
    });
    expect(rightResult.ok).toBe(true);
    expect((rightResult.operation as { alignment?: string }).alignment).toBe('right');

    const leftResult = validateOperationStrict({
      action: 'table',
      target: 'tbl-1',
      operation: 'setTableAlignment',
      align: 'Left',
    });
    expect(leftResult.ok).toBe(true);
    expect((leftResult.operation as { align?: string }).align).toBe('left');

    for (const invalid of ['MIDDLE', 'top']) {
      expect(() =>
        validateOperationStrict({
          action: 'table',
          target: 'tbl-1',
          operation: 'setTableAlignment',
          align: invalid,
        })
      ).toThrow(/^E_VALIDATION:/);

      expect(() =>
        validateOperationStrict({
          action: 'table',
          target: 'tbl-1',
          operation: 'setTableAlignment',
          alignment: invalid,
        })
      ).toThrow(/^E_VALIDATION:/);
    }
  });

  it('executes table alignment via engine with uppercase and mixed-case values and sets w:jc correctly', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;

    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'CENTER',
    });

    let tblPr = directChildren(tblEl, 'w:tblPr')[0];
    let jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc).toBeDefined();
    expect(jc.getAttribute('w:val')).toBe('center');

    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      alignment: 'RIGHT',
    });

    tblPr = directChildren(tblEl, 'w:tblPr')[0];
    jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc).toBeDefined();
    expect(jc.getAttribute('w:val')).toBe('right');

    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'Left',
    });

    tblPr = directChildren(tblEl, 'w:tblPr')[0];
    jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc).toBeDefined();
    expect(jc.getAttribute('w:val')).toBe('left');
  });

  it('updates existing w:jc in-place without duplicating elements', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;

    // Apply alignment 3 times
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'left',
    } as never);

    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'center',
    } as never);

    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'right',
    } as never);

    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    const jcElements = directChildren(tblPr, 'w:jc');
    expect(jcElements).toHaveLength(1);
    expect(jcElements[0].getAttribute('w:val')).toBe('right');
  });

  it('throws E_VALIDATION when target table node id is missing', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    await insertTable(engine, lastNodeId(ast));

    // When called directly without target
    try {
      await executeTable(
        {
          action: 'table',
          operation: 'setTableAlignment',
          align: 'center',
        },
        ast
      );
      expect.unreachable('Should have thrown for missing target');
    } catch (err: unknown) {
      const error = err as { code?: string; message?: string; remedy?: string };
      expect(error.code).toBe('E_VALIDATION');
      expect(error.message).toContain('Table operation "setTableAlignment" requires a target table node id.');
      expect(error.remedy).toContain('Pass target: "<tableNodeId>" in the operation payload.');
    }

    // When called through operation engine without target
    await expect(
      engine.execute({
        action: 'table',
        operation: 'setTableAlignment',
        align: 'center',
      } as never)
    ).rejects.toThrow('Table operation "setTableAlignment" requires a target table node id.');
  });

  it('throws E_VALIDATION when target is missing even on a document with no tables', async () => {
    const { ast } = await freshDoc();

    try {
      await executeTable(
        {
          action: 'table',
          operation: 'setTableAlignment',
          align: 'center',
        },
        ast
      );
      expect.unreachable('Should have thrown for missing target');
    } catch (err: unknown) {
      const error = err as { code?: string; message?: string; remedy?: string };
      expect(error.code).toBe('E_VALIDATION');
      expect(error.message).toContain('Table operation "setTableAlignment" requires a target table node id.');
      expect(error.remedy).toContain('Pass target: "<tableNodeId>" in the operation payload.');
    }
  });

  it('throws E_VALIDATION when alignment value is invalid', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));

    const invalidValues = ['justify', 'start', 'end', '', undefined, 123];

    for (const invalid of invalidValues) {
      try {
        await executeTable(
          {
            action: 'table',
            target: tableId,
            operation: 'setTableAlignment',
            align: invalid as unknown as string,
          },
          ast
        );
        expect.unreachable(`Should have thrown for invalid alignment: ${invalid}`);
      } catch (err: unknown) {
        const error = err as { code?: string; message?: string; remedy?: string };
        expect(error.code).toBe('E_VALIDATION');
        expect(error.message).toContain(`Invalid table alignment "${invalid}". Expected "left", "center", or "right".`);
        expect(error.remedy).toContain('Pass align: "left" | "center" | "right" in the operation payload.');
      }
    }
  });

  it('asserts w:jc position conforms to schema order (after w:tblW, before borders/spacing)', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;
    const doc = tblEl.ownerDocument;

    // Pre-populate tblPr with elements that should surround w:jc in TBLPR_CHILD_ORDER
    let tblPr = directChildren(tblEl, 'w:tblPr')[0];
    if (!tblPr) {
      tblPr = createWordElement(doc, 'w:tblPr');
      tblEl.insertBefore(tblPr, tblEl.firstChild);
    }

    // Add w:tblBorders, w:tblCellMar, and w:tblLook after tblW
    const tblBorders = createWordElement(doc, 'w:tblBorders');
    tblPr.appendChild(tblBorders);

    const tblCellMar = createWordElement(doc, 'w:tblCellMar');
    tblPr.appendChild(tblCellMar);

    const tblLook = createWordElement(doc, 'w:tblLook');
    tblPr.appendChild(tblLook);

    // Apply table alignment
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'center',
    } as never);

    // Collect child tag names of tblPr in order
    const childTags: string[] = [];
    for (let i = 0; i < tblPr.childNodes.length; i++) {
      const child = tblPr.childNodes[i];
      if (child.nodeType === 1) {
        childTags.push((child as Element).tagName);
      }
    }

    const tblWIndex = childTags.indexOf('w:tblW');
    const jcIndex = childTags.indexOf('w:jc');
    const tblBordersIndex = childTags.indexOf('w:tblBorders');
    const tblCellMarIndex = childTags.indexOf('w:tblCellMar');
    const tblLookIndex = childTags.indexOf('w:tblLook');

    expect(jcIndex).toBeGreaterThan(-1);

    if (tblWIndex !== -1) {
      expect(jcIndex).toBeGreaterThan(tblWIndex);
    }
    if (tblBordersIndex !== -1) {
      expect(jcIndex).toBeLessThan(tblBordersIndex);
    }
    if (tblCellMarIndex !== -1) {
      expect(jcIndex).toBeLessThan(tblCellMarIndex);
    }
    if (tblLookIndex !== -1) {
      expect(jcIndex).toBeLessThan(tblLookIndex);
    }

    // Verify all child elements follow exact TBLPR_CHILD_ORDER sequence
    const orderIndices = childTags.map((tag) => TBLPR_CHILD_ORDER.indexOf(tag));
    for (let i = 0; i < orderIndices.length - 1; i++) {
      if (orderIndices[i] !== -1 && orderIndices[i + 1] !== -1) {
        expect(orderIndices[i]).toBeLessThanOrEqual(orderIndices[i + 1]);
      }
    }
  });

  it('creates w:tblPr if not already present on target table', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const table = tableNodeOf(ast);
    const tblEl = table.xml as Element;

    // Remove w:tblPr completely to test creation branch
    const existingTblPr = directChildren(tblEl, 'w:tblPr')[0];
    if (existingTblPr) {
      tblEl.removeChild(existingTblPr);
    }
    expect(directChildren(tblEl, 'w:tblPr')).toHaveLength(0);

    // Apply alignment
    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'center',
    } as never);

    const newTblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(newTblPr).toBeDefined();
    expect(tblEl.firstChild).toBe(newTblPr);
    const jc = directChildren(newTblPr, 'w:jc')[0];
    expect(jc).toBeDefined();
    expect(jc.getAttribute('w:val')).toBe('center');
  });

  it('survives document save and reload round-trip with XML integrity intact', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));

    await engine.execute({
      action: 'table',
      target: tableId,
      operation: 'setTableAlignment',
      align: 'center',
    } as never);

    // Save document to buffer (runs internal zip & XML integrity validation)
    const buffer = parser.rezip();
    expect(buffer).toBeInstanceOf(Buffer);
    expect(buffer.length).toBeGreaterThan(0);

    // Reload from buffer
    const reparsed = new DocxParser(buffer);
    const ast2 = await reparsed.getAST();
    const reloadedTable = ast2.sections[0].nodes.find((n) => n.type === 'table');
    expect(reloadedTable).toBeDefined();

    const tblEl = reloadedTable!.xml as Element;
    const tblPr = directChildren(tblEl, 'w:tblPr')[0];
    expect(tblPr).toBeDefined();
    const jc = directChildren(tblPr, 'w:jc')[0];
    expect(jc).toBeDefined();
    expect(jc.getAttribute('w:val')).toBe('center');

    // Full OOXML schema validation
    const validation = await reparsed.validateOOXML();
    expect(validation.ok).toBe(true);
  });

  it('throws E_VALIDATION when targeting a paragraph node and leaves paragraph XML untouched', async () => {
    const { parser, ast } = await freshDoc();
    const engine = new OperationEngine(ast, parser);
    const tableId = await insertTable(engine, lastNodeId(ast));
    const paragraphNode = ast.sections[0].nodes.find((n) => n.type === 'paragraph');
    expect(paragraphNode).toBeDefined();
    const paragraphId = paragraphNode!.id;
    const initialXml = new XMLSerializer().serializeToString(paragraphNode!.xml);

    // Direct executeTable call fails closed with typed E_VALIDATION
    await expect(
      executeTable(
        {
          action: 'table',
          target: paragraphId,
          operation: 'setTableAlignment',
          align: 'center',
        },
        ast
      )
    ).rejects.toMatchObject({
      code: 'E_VALIDATION',
      message: expect.stringContaining(`Target '${paragraphId}' is a paragraph (<p>), but table operations require a table target.`),
      remedy: expect.stringContaining(`Pass a valid table target id (e.g. ${tableId}) from inspect_document.`),
    });

    // Verify paragraph XML is completely untouched
    const afterDirectXml = new XMLSerializer().serializeToString(paragraphNode!.xml);
    expect(afterDirectXml).toBe(initialXml);
    expect(directChildren(paragraphNode!.xml, 'w:tblPr')).toHaveLength(0);
    expect(directChildren(paragraphNode!.xml, 'w:tblGrid')).toHaveLength(0);

    // Engine execute call fails closed
    await expect(
      engine.execute({
        action: 'table',
        target: paragraphId,
        operation: 'setTableAlignment',
        align: 'center',
      } as never)
    ).rejects.toThrow(`Target '${paragraphId}' is a paragraph (<p>), but table operations require a table target.`);

    // Verify paragraph XML remains completely untouched after engine execution
    const afterEngineXml = new XMLSerializer().serializeToString(paragraphNode!.xml);
    expect(afterEngineXml).toBe(initialXml);
    expect(directChildren(paragraphNode!.xml, 'w:tblPr')).toHaveLength(0);
    expect(directChildren(paragraphNode!.xml, 'w:tblGrid')).toHaveLength(0);

    // Engine executeBatch call fails closed with rollback and returns failure
    const batchResult = await engine.executeBatch({
      operations: [
        {
          action: 'table',
          operation: 'setTableAlignment',
          target: paragraphId,
          params: { align: 'center' },
        },
      ],
    });
    expect(batchResult.success).toBe(false);
    expect(batchResult.appliedCount).toBe(0);
    expect(batchResult.errors[0].error).toContain(`Target '${paragraphId}' is a paragraph (<p>), but table operations require a table target.`);
    expect(batchResult.errors[0].remedy).toContain('Pass a valid table target id');
    expect(batchResult.errors[0].remedy).toContain('from inspect_document.');

    const afterBatchXml = new XMLSerializer().serializeToString(paragraphNode!.xml);
    expect(afterBatchXml).toBe(initialXml);
    expect(directChildren(paragraphNode!.xml, 'w:tblPr')).toHaveLength(0);
    expect(directChildren(paragraphNode!.xml, 'w:tblGrid')).toHaveLength(0);

    // Full OOXML validation remains valid
    const validation = await parser.validateOOXML();
    expect(validation.ok).toBe(true);
  });
});
