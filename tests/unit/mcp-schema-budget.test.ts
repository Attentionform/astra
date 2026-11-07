import { describe, expect, it } from 'vitest';
import { getToolDefinitions, compactToolSchema } from '../../packages/mcp-server/src/lib/mcp/schema-dump';

describe('MCP tool schema budget', () => {
  it('keeps the complete compact tool catalog under 90,000 characters', () => {
    const tools = getToolDefinitions().map(tool => ({
      name:tool.name, description:tool.description,
      parameters:compactToolSchema(tool.schema),
    }));
    expect(JSON.stringify(tools).length).toBeLessThan(90_000);
    expect(JSON.stringify(tools.find(tool => tool.name === 'execute_batch')?.parameters)).toContain('$ref');
    for (const tool of tools) {
      const root = tool.parameters;
      const walk = (value: unknown): void => {
        if (!value || typeof value !== 'object') return;
        const object = value as Record<string, unknown>;
        if (typeof object.$ref === 'string') {
          expect(object.$ref.startsWith('#/')).toBe(true);
          let target: unknown = root;
          for (const part of object.$ref.slice(2).split('/')) {
            target = (target as Record<string, unknown> | undefined)?.[part.replace(/~1/g, '/').replace(/~0/g, '~')];
          }
          expect(target, `${tool.name}: ${object.$ref}`).toBeDefined();
        }
        for (const child of Object.values(object)) walk(child);
      };
      walk(root);
    }
  });
});
