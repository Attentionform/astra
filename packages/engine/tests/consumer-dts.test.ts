/**
 * @attentionform/astra-engine — consumer typecheck proof (proper declaration builds).
 *
 * Proves the shipped engine dist typings resolve for a real TS consumer:
 * a scratch file importing `@attentionform/astra-engine` BY NAME typechecks clean with
 * real `tsc --noEmit` (types entry from package.json, no path aliases).
 * Dist-dependent cases skip when dist/ was not built (same runIf pattern
 * as the agent packaging tests).
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

describe('engine consumer typecheck (npm distribution)', () => {
  it.runIf(distBuilt)('declared types entry exists and re-exports the engine surface', () => {
    expect(pkg.name).toBe('@attentionform/astra-engine');
    const dts = fs.readFileSync(distTypes, 'utf8');
    expect(dts).toMatch(/export \* from '\.\/document-manager'/);
    expect(dts).toMatch(/export \* from '\.\/parser'/);
  });

  it.runIf(distBuilt)('scratch consumer importing @attentionform/astra-engine typechecks clean', () => {
    expect(() =>
      typecheckConsumer(
        [
          "import { DocumentManager, DocxParser, OperationEngine } from '@attentionform/astra-engine';",
          'export type AnyManager = DocumentManager;',
          'export type AnyParser = DocxParser;',
          'export function describeEngine(): string {',
          '  void DocumentManager;',
          '  void DocxParser;',
          '  void OperationEngine;',
          "  return 'ok';",
          '}',
          '',
        ].join('\n')
      )
    ).not.toThrow();
  });
});
