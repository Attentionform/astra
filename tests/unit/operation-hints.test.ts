import { describe, it, expect } from 'vitest';
import {
  BatchOperationsSchema,
  closestNames,
  createBlankDocxBuffer,
  DocxParser,
  operationIssueRemedy,
} from '@attentionform/astra-engine';
import { executeStyle } from '../../packages/engine/src/operations/style';
import { handleExecuteBatch } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';

/**
 * The schema gate says `Invalid discriminator value` for an unknown action and `Unrecognized key` for an
 * unknown key. The remedy now says what to write instead and lists what the action accepts.
 */
function remedyFor(operation: Record<string, unknown>, pointer?: string): string | undefined {
  const operations = [operation];
  const parsed = BatchOperationsSchema.safeParse({ operations });
  expect(parsed.success).toBe(false);
  if (parsed.success) return undefined;
  return operationIssueRemedy(operations, parsed.error.issues, pointer ? { pointer } : {});
}

describe('closestNames', () => {
  const actions = ['insert', 'replace', 'delete', 'style', 'hyperlink'];
  const words = { remove: 'delete', link: 'hyperlink' };

  it('matches without regard to case or separators', () => {
    expect(closestNames('Delete', actions)).toEqual(['delete']);
    expect(closestNames('hyper_link', actions)).toEqual(['hyperlink']);
  });

  it('uses a known word before the edit distance', () => {
    expect(closestNames('remove', actions, words)).toEqual(['delete']);
  });

  it('accepts a small slip and gives nothing for a name that is far from all', () => {
    expect(closestNames('replce', actions)).toEqual(['replace']);
    expect(closestNames('zzzzzzzz', actions)).toEqual([]);
    expect(closestNames('', actions)).toEqual([]);
  });
});

describe('unknown action', () => {
  it('"remove" suggests "delete" and lists the actions', () => {
    const remedy = remedyFor({ action: 'remove', target: 'x' }, 'Run `astra batch --schema` for an example of each action.')!;
    expect(remedy).toContain('operation 0: "remove" is not an action.');
    expect(remedy).toContain('Did you mean "delete"?');
    expect(remedy).toMatch(/Actions: insert, replace, delete, style,/);
    expect(remedy.endsWith('Run `astra batch --schema` for an example of each action.')).toBe(true);
  });

  it('a wrong case is named', () => {
    expect(remedyFor({ action: 'Style', target: 'x' })).toContain('Did you mean "style"?');
  });

  it('an action that is not close to any gets the list and no guess', () => {
    const remedy = remedyFor({ action: 'zzzzzzzz' })!;
    expect(remedy).not.toContain('Did you mean');
    expect(remedy).toContain('Actions: insert');
  });
});

describe('unknown key', () => {
  it('"spaceAfter" on style suggests "spacingAfter" and lists the keys of style', () => {
    const remedy = remedyFor({ action: 'style', target: 'x', spaceAfter: 120 })!;
    expect(remedy).toContain('"spaceAfter" is not a key of "style". Did you mean "spacingAfter"?');
    expect(remedy).toMatch(/Keys of "style": target,.*spacingAfter/);
  });

  it('"spaceAfter" inside params suggests "spacingAfter" on execute', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const target = ast.sections[0].nodes[0].id;
    expect(() => executeStyle({ action: 'style', target, params: { spaceAfter: 120 } } as never, ast)).toThrow(
      /Did you mean "spacingAfter"/
    );
  });

  it('"spaceAfter" inside properties suggests "spacingAfter" on execute', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const target = ast.sections[0].nodes[0].id;
    expect(() => executeStyle({ action: 'style', target, properties: { spaceAfter: 120 } } as never, ast)).toThrow(
      /Did you mean "spacingAfter"/
    );
  });

  it('names every unknown key, and guesses only where it can', () => {
    const remedy = remedyFor({ action: 'style', target: 'x', boldd: true, zzzz: 1 })!;
    expect(remedy).toContain('"boldd" is not a key of "style". Did you mean "bold"?');
    expect(remedy).toContain('"zzzz" is not a key of "style".');
    expect(remedy).not.toContain('"zzzz" is not a key of "style". Did you mean');
  });

  it('a long key list is cut', () => {
    const remedy = remedyFor({ action: 'table', target: 'x', operation: 'addRow', nonsense: 1 })!;
    expect(remedy).toContain('Keys of "table":');
  });
});

describe('unknown operation name', () => {
  it('"addRows" on table suggests "addRow"', () => {
    const remedy = remedyFor({ action: 'table', operation: 'addRows', target: 'x' })!;
    expect(remedy).toBe('operation 0: "addRows" is not an operation of "table". Did you mean "addRow"?');
  });
});

describe('issues that the hints do not cover', () => {
  it('give no remedy, so the caller keeps its own', () => {
    const operations = [{ action: 'delete' }];
    const parsed = BatchOperationsSchema.safeParse({ operations });
    expect(parsed.success).toBe(false);
    if (!parsed.success) expect(operationIssueRemedy(operations, parsed.error.issues)).toBeUndefined();
  });

  it('an operations value that is not a list gives none and does not fail', () => {
    expect(operationIssueRemedy(undefined, [{ code: 'invalid_union', path: ['operations', 0, 'action'] }])).toBeUndefined();
  });
});

describe('MCP execute_batch', () => {
  it('gives the same hints', async () => {
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(createBlankDocxBuffer('blank'), 'hints.docx');
    const target = handle.ast.sections[0].nodes[0].id;
    const out = await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'remove', target }] });
    expect(out.ok).toBe(false);
    const text = JSON.stringify(out);
    expect(text).toContain('Did you mean \\"delete\\"?');
    expect(text).toContain('execute_batch input schema');
  });
});
