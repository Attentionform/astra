/**
 * Error text for library callers. The remedies keep naming the MCP tools (agents act on those names);
 * `libraryRemedy` adds the library equivalent next to each tool and never replaces the original.
 */
import { describe, expect, it } from 'vitest';
import { DocxParser, OperationEngine, libraryRemedy, toError, MCP_TOOL_LIBRARY_EQUIVALENTS } from '../src';

async function failedBatch(operation: Record<string, unknown>) {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  return engine.executeBatch({ operations: [operation as never] });
}

describe('libraryRemedy', () => {
  it('names the library equivalent next to the MCP tool and keeps the tool name', () => {
    const text = 're-read ids via inspect_document, then retry with execute_batch';
    const out = libraryRemedy(text);
    expect(out).toContain('inspect_document (library: parser.getStructuralSummary() or parser.getAST())');
    expect(out).toContain('execute_batch (library: OperationEngine.executeBatch())');
  });
  it('names each tool once and returns undefined when no tool is named', () => {
    expect(libraryRemedy('inspect_document, then inspect_document again')?.match(/\(library:/g)).toHaveLength(1);
    expect(libraryRemedy('check the path and retry')).toBeUndefined();
    expect(libraryRemedy(undefined)).toBeUndefined();
  });
  it('every mapped tool is replaced by its own equivalent', () => {
    for (const [tool, equivalent] of Object.entries(MCP_TOOL_LIBRARY_EQUIVALENTS)) {
      expect(libraryRemedy(`use ${tool}`)).toBe(`use ${tool} (library: ${equivalent})`);
    }
  });
});

describe('toError', () => {
  it('keeps the MCP remedy text and adds libraryRemedy', () => {
    const err = toError('E_SESSION', 'no session');
    expect(err.remedy).toContain('open_document');
    expect(err.message).toContain('open_document');
    expect(err.libraryRemedy).toContain('open_document (library: new DocxParser(buffer))');
  });
  it('adds nothing when the remedy names no tool', () => {
    expect(toError('E_IO', 'unreadable').libraryRemedy).toBeUndefined();
  });
});

describe('a failed batch', () => {
  it("tells a caller who inserts type 'heading' what to use instead", async () => {
    const result = await failedBatch({ action: 'insert', type: 'heading', anchor: 'end', position: 'after', text: 'Title' });
    expect(result.success).toBe(false);
    const text = result.errors[0].error;
    expect(text).toContain("use type 'paragraph' with style 'Heading1'");
    // The original wording for agents is still there.
    expect(text).toContain('inspect_document');
  });
  it('keeps the MCP tool names in remedy and hint and adds libraryRemedy beside them', async () => {
    const result = await failedBatch({ action: 'replace', target: 'p_does_not_exist', text: 'x' });
    expect(result.success).toBe(false);
    const entry = result.errors[0];
    const agentText = `${entry.remedy ?? ''} ${entry.hint ?? ''}`;
    expect(agentText).toContain('inspect_document');
    expect(entry.libraryRemedy).toContain('inspect_document (library:');
  });
});
