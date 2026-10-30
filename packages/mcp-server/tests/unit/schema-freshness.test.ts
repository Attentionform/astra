import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { getServedToolNames, SERVER_VERSION } from '../../src/index';
import {
  buildStamp,
  checkSchemasInDir,
  formatMismatchWarning,
  generateSchemasToDir,
  getToolDefinitions,
} from '../../src/lib/mcp/schema-dump';

function freshTmpDir(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  return dir;
}

describe('schema freshness chain', () => {
  it('generator covers every served tool with no silent gaps', () => {
    const served = [...getServedToolNames()].sort();
    const generated = getToolDefinitions().map((t) => t.name).sort();
    expect(generated).toEqual(served);
    expect(generated).toContain('compose_document');
    expect(generated).toHaveLength(served.length);
  });

  it('generated files carry the installed server version stamp', () => {
    const dir = freshTmpDir('astra-schemas-stamp-');
    const result = generateSchemasToDir(dir, SERVER_VERSION);
    expect(result.count).toBe(getServedToolNames().length);
    const expected = buildStamp(SERVER_VERSION);
    for (const tool of getServedToolNames()) {
      const raw = JSON.parse(fs.readFileSync(path.join(dir, `${tool}.json`), 'utf8')) as Record<string, unknown>;
      expect(raw['x-generated-from']).toBe(expected);
      expect(raw['name']).toBe(tool);
      expect(typeof raw['description']).toBe('string');
      expect((raw['description'] as string).length).toBeGreaterThan(0);
      const params = raw['parameters'] as { type?: string } | undefined;
      expect(params?.type).toBe('object');
    }
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('fresh copies pass the version check with zero mismatches', () => {
    const dir = freshTmpDir('astra-schemas-fresh-');
    generateSchemasToDir(dir, SERVER_VERSION);
    const verdict = checkSchemasInDir(dir, SERVER_VERSION);
    expect(verdict.ok).toBe(true);
    expect(verdict.mismatches).toHaveLength(0);
    expect(verdict.checked).toBe(getServedToolNames().length);
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('a stale stamp fails the check with a refresh warning', () => {
    const dir = freshTmpDir('astra-schemas-stale-');
    generateSchemasToDir(dir, SERVER_VERSION);
    const target = path.join(dir, 'compose_document.json');
    const raw = JSON.parse(fs.readFileSync(target, 'utf8')) as Record<string, unknown>;
    raw['x-generated-from'] = 'astra-mcp-server@0.0.0-stale';
    fs.writeFileSync(target, JSON.stringify(raw, null, 2) + '\n', 'utf8');
    const verdict = checkSchemasInDir(dir, SERVER_VERSION);
    expect(verdict.ok).toBe(false);
    expect(verdict.mismatches.length).toBeGreaterThanOrEqual(1);
    const hit = verdict.mismatches.find((m) => m.file.endsWith('compose_document.json'));
    expect(hit).toBeDefined();
    expect(hit?.stamped).toBe('astra-mcp-server@0.0.0-stale');
    expect(hit?.expected).toBe(buildStamp(SERVER_VERSION));
    const warning = formatMismatchWarning(verdict.mismatches);
    expect(warning).toContain('--dump-schemas');
    expect(warning).toContain('compose_document.json');
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('a missing copy fails the check instead of passing silently', () => {
    const dir = freshTmpDir('astra-schemas-missing-');
    generateSchemasToDir(dir, SERVER_VERSION);
    fs.unlinkSync(path.join(dir, 'open_document.json'));
    const verdict = checkSchemasInDir(dir, SERVER_VERSION);
    expect(verdict.ok).toBe(false);
    expect(verdict.mismatches.some((m) => m.file.endsWith('open_document.json'))).toBe(true);
    fs.rmSync(dir, { recursive: true, force: true });
  });
});
