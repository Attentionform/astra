import * as path from 'path';
import { getServedToolNames, SERVER_VERSION } from '../packages/mcp-server/src/index';
import { generateSchemasToDir, getToolDefinitions } from '../packages/mcp-server/src/lib/mcp/schema-dump';

function assertFullCoverage(): void {
  const served = [...getServedToolNames()].sort();
  const generated = getToolDefinitions().map((t) => t.name).sort();
  const missing = served.filter((n) => !generated.includes(n));
  const extra = generated.filter((n) => !served.includes(n));
  if (missing.length > 0 || extra.length > 0) {
    throw new Error(
      `[generate_mcp_schemas] tool coverage drift: missing [${missing.join(', ')}], extra [${extra.join(', ')}]. ` +
        'Generator must byte-cover every served tool.'
    );
  }
}

function resolveTarget(): string {
  const raw = process.argv[2];
  if (raw) return path.resolve(raw);
  return path.resolve(__dirname, '../schemas/mcp');
}

function run(): void {
  assertFullCoverage();
  const target = resolveTarget();
  const result = generateSchemasToDir(target, SERVER_VERSION);
  process.stdout.write(`[generate_mcp_schemas] Generating schemas for ${result.count} tools into: ${target}\n`);
  process.stdout.write(`[generate_mcp_schemas] Successfully wrote ${result.count} schema files.\n`);
}

const entry = process.argv[1] ?? '';
if (entry.endsWith('generate_mcp_schemas.ts') || entry.endsWith('generate_mcp_schemas.js')) {
  run();
}

export { assertFullCoverage, resolveTarget, run };
