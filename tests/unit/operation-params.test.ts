/**
 * Free-form `params` are checked, not ignored.
 *
 * `{ action:'table', params:{ op:'updateCell', bold:true } }` used to come back ok:true and change
 * nothing, because the schema let any key through and the executor read only the keys it knew.
 * Now every key an operation carries must be one it reads, and a bad batch is refused whole.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executeTable,
  validateOperationStrict,
  validateOperationsStrict,
  BatchOperationsSchema,
  OperationSchema,
  NESTED_PARAM_KEYS,
} from '@attentionform/astra-engine';
import { TABLE_SUBOP_KEYS, unknownParamsError } from '../../packages/engine/src/operation-params';

const updateCell = { action: 'table', target: 't_001', params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'x' } };

describe('table params are checked against the sub-operation', () => {
  it('rejects bold on updateCell and names the key and what is allowed', () => {
    const op = { ...updateCell, params: { ...updateCell.params, bold: true } };
    expect(() => validateOperationStrict(op)).toThrow(
      /unknown parameter "bold" for table updateCell; allowed: rowIndex, row, colIndex, col, columnIndex, text/,
    );
  });

  it('rejects a flat key the schema knows but this sub-operation does not read', () => {
    const op = { action: 'table', target: 't_001', op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'x', fill: 'FF0000' };
    expect(() => validateOperationStrict(op)).toThrow(/unknown parameter "fill" for table updateCell/);
  });

  it('lists every unknown key, not just the first', () => {
    const op = { ...updateCell, params: { ...updateCell.params, bold: true, italic: true } };
    expect(() => validateOperationStrict(op)).toThrow(/"bold", "italic"/);
  });

  it('still accepts the documented shapes, flat and nested', () => {
    const ok: unknown[] = [
      updateCell,
      { action: 'table', target: 't_001', op: 'updateCell', row: 2, col: 3, text: 'y' },
      { action: 'table', target: 't_001', params: { op: 'setBorder', border: 'all', color: 'FF0000', sz: 8, val: 'single' } },
      { action: 'table', target: 't_001', op: 'setShading', fill: 'F7F9FC', headerFill: '1E3A8A', zebra: true, textColor: 'FFFFFF' },
      { action: 'table', target: 't_001', params: { op: 'addRow', afterIndex: 1, cells: ['a', 'b'] } },
      { action: 'table', target: 't_001', op: 'setCantSplit' },
      { action: 'table', target: 't_001', op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'z', trackChanges: true, expectedVersion: 1 },
    ];
    for (const op of ok) expect(() => validateOperationStrict(op), JSON.stringify(op)).not.toThrow();
  });

  it('refuses a batch whole when one operation is bad, naming which', () => {
    const bad = { ...updateCell, params: { ...updateCell.params, bold: true } };
    expect(() => validateOperationsStrict([updateCell, bad])).toThrow(/operations\[1\].*unknown parameter "bold"/);
  });

  it('is enforced by the batch schema the CLI and the MCP server both parse with', () => {
    const bad = { ...updateCell, params: { ...updateCell.params, bold: true } };
    const result = BatchOperationsSchema.safeParse({ operations: [updateCell, bad] });
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.issues[0].path).toEqual(['operations', 1]);
      expect(result.error.issues[0].message).toMatch(/unknown parameter "bold" for table updateCell/);
    }
    expect(BatchOperationsSchema.safeParse({ operations: [updateCell] }).success).toBe(true);
  });

  it('has a key list for every table sub-operation in the schema, and no extra', () => {
    const table = OperationSchema.options.find((o) => (o.shape.action as { value: string }).value === 'table')!;
    const subOps = ((table.shape as Record<string, any>).operation.unwrap().options as string[]).slice().sort();
    expect(Object.keys(TABLE_SUBOP_KEYS).sort()).toEqual(subOps);
  });
});

describe('every operation with free-form params is covered', () => {
  const freeForm = OperationSchema.options
    .map((o) => ({ action: (o.shape.action as { value: string }).value, params: (o.shape as Record<string, any>).params }))
    .filter(({ params }) => params?._zod?.def?.type === 'optional' && params._zod.def.innerType._zod.def.type === 'record')
    .map(({ action }) => action);

  it('finds the eleven operations the audit named', () => {
    expect(freeForm.sort()).toEqual(
      ['callout', 'chart', 'clearFormat', 'delete', 'equation', 'field', 'list', 'move', 'replace', 'section', 'table'].sort(),
    );
  });

  it('has a policy for each one, and rejects a key it does not read', () => {
    for (const action of freeForm.filter((a) => a !== 'table')) {
      expect(NESTED_PARAM_KEYS[action], `${action} has no nested key list`).toBeDefined();
      const err = unknownParamsError({ action, params: { __nope: 1 } }, NESTED_PARAM_KEYS);
      expect(err?.message, action).toMatch(/unknown parameter "__nope"/);
    }
  });

  it('accepts, nested, every key the same operation accepts flat', () => {
    for (const option of OperationSchema.options) {
      const shape = option.shape as Record<string, unknown>;
      const action = (shape.action as { value: string }).value;
      if (!(action in NESTED_PARAM_KEYS)) continue;
      for (const key of Object.keys(shape).filter((k) => k !== 'action' && k !== 'params')) {
        expect(unknownParamsError({ action, params: { [key]: 1 } }, NESTED_PARAM_KEYS), `${action}.${key}`).toBeNull();
      }
    }
  });

  it('keeps the nested-only keys the executors read', () => {
    expect(unknownParamsError({ action: 'replace', params: { resetStyle: true } }, NESTED_PARAM_KEYS)).toBeNull();
    expect(unknownParamsError({ action: 'delete', params: { trackChanges: true } }, NESTED_PARAM_KEYS)).toBeNull();
  });

  it('rejects an invented key on replace through the public validator', () => {
    expect(() => validateOperationStrict({ action: 'replace', target: 'p_1', text: 'x', params: { bold: true } })).toThrow(
      /unknown parameter "bold" for replace/,
    );
  });
});

describe('flat keys reach the sub-operations that used to read only params', () => {
  async function freshTable() {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    await executeInsert(
      { action: 'insert', type: 'table', anchor: ast.sections[0].nodes[0].id, position: 'after', headers: ['A', 'B'], rows: [['1', '2']] } as any,
      ast,
      parser,
    );
    const tableId = ast.sections[0].nodes.find((n: any) => n.type === 'table').id as string;
    const xml = () => new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
    return { ast, tableId, xml };
  }

  it('setBorder honours a flat border, color and sz', async () => {
    const { ast, tableId, xml } = await freshTable();
    await executeTable({ action: 'table', target: tableId, op: 'setBorder', border: 'top', color: 'FF0000', sz: 12 } as any, ast);
    const top = xml().match(/<w:tblBorders>[\s\S]*?<w:top ([^>]*)\/>/)![1];
    expect(top).toContain('w:sz="12"');
    expect(top).toContain('w:color="FF0000"');
  });

  it('setBorder refuses a size that is not a whole number of eighths of a point', async () => {
    const { ast, tableId } = await freshTable();
    await expect(
      executeTable({ action: 'table', target: tableId, params: { op: 'setBorder', sz: '1pt' } } as any, ast),
    ).rejects.toThrow(/eighths of a point/);
  });

  it('setVerticalAlign honours a flat vAlign', async () => {
    const { ast, tableId, xml } = await freshTable();
    await executeTable({ action: 'table', target: tableId, op: 'setVerticalAlign', vAlign: 'bottom' } as any, ast);
    expect(xml()).toContain('<w:vAlign w:val="bottom"/>');
  });

  it('addColumn honours a flat afterIndex and width', async () => {
    const { ast, tableId, xml } = await freshTable();
    await executeTable({ action: 'table', target: tableId, op: 'addColumn', afterIndex: 1, width: 3333 } as any, ast);
    expect(xml()).toContain('w:w="3333"');
  });
});
