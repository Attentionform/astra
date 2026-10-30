/**
 * @attentionform/astra-mcp-server — consumer typecheck proof (proper declaration builds).
 *
 * Proves the shipped dist typings are real tsc declarations (not the old
 * inline-string stub) and resolve for a real TS consumer importing
 * `@attentionform/astra-mcp-server` BY NAME with real `tsc --noEmit`.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { execFileSync } from 'child_process';

const pkgDir = path.join(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(pkgDir, 'package.json'), 'utf8')) as {
  name: string;
  main: string;
  types: string;
};
const distTypes = path.join(pkgDir, pkg.types);
const distMain = path.join(pkgDir, pkg.main);
const distBuilt = fs.existsSync(distTypes) && fs.existsSync(distMain);
const repoRoot = path.join(pkgDir, '..', '..');
const tscBin = path.join(repoRoot, 'node_modules', 'typescript', 'lib', 'tsc.js');

function typecheckConsumer(source: string): void {
  const dir = fs.mkdtempSync(path.join(pkgDir, '.consumer-check-'));
  try {
    fs.writeFileSync(path.join(dir, 'consumer.ts'), source);
    fs.writeFileSync(
      path.join(dir, 'tsconfig.json'),
      JSON.stringify({
        compilerOptions: {
          target: 'ES2022',
          module: 'NodeNext',
          moduleResolution: 'NodeNext',
          strict: true,
          skipLibCheck: true,
          types: ['node'],
          noEmit: true,
        },
        include: ['consumer.ts'],
      })
    );
    try {
      execFileSync(process.execPath, [tscBin, '-p', path.join(dir, 'tsconfig.json')], {
        encoding: 'utf8',
        timeout: 120000,
      });
    } catch (err) {
      const e = err as { stdout?: unknown; stderr?: unknown; message?: string };
      throw new Error(
        `consumer typecheck failed:\n${String(e.stdout ?? '')}\n${String(e.stderr ?? '')}\n${String(e.message ?? e)}`
      );
    }
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

describe('mcp-server consumer typecheck (npm distribution)', () => {
  it.runIf(distBuilt)('dist/index.d.ts is full tsc surface (not the inline-string stub)', () => {
    expect(pkg.name).toBe('@attentionform/astra-mcp-server');
    const dts = fs.readFileSync(distTypes, 'utf8');
    // The old build wrote a 5-line stub with only SERVER_NAME/SERVER_VERSION/
    // createAstraMcpServer/main. Real declarations cover the tool surface too.
    expect(dts).toMatch(/\bLEAN6_TOOL_NAMES\b/);
    expect(dts).toMatch(/\bcreateAstraMcpServer\b/);
    expect(dts.split('\n').length).toBeGreaterThan(10);
  });

  it.runIf(distBuilt)('scratch consumer importing @attentionform/astra-mcp-server typechecks clean', () => {
    expect(() =>
      typecheckConsumer(
        [
          "import { createAstraMcpServer, SERVER_NAME, LEAN6_TOOL_NAMES } from '@attentionform/astra-mcp-server';",
          'export function describeServer(): string {',
          '  void createAstraMcpServer;',
          '  void LEAN6_TOOL_NAMES;',
          '  return SERVER_NAME;',
          '}',
          '',
        ].join('\n')
      )
    ).not.toThrow();
  });
});
