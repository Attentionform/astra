/**
 * @attentionform/astra — consumer typecheck proof (proper declaration builds).
 *
 * Proves the shipped dist typings are real tsc declarations and resolve for
 * a real TS consumer
 * importing `@attentionform/astra` BY NAME with real `tsc --noEmit`.
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

describe('cli consumer typecheck (npm distribution)', () => {
  it.runIf(distBuilt)('declared types entry is real cli tsc', () => {
    expect(pkg.name).toBe('@attentionform/astra');
    const distDts = fs.readFileSync(distTypes, 'utf8');
    expect(distDts).toMatch(/\brun\b/);
    // cli.d.ts is a barrel: CLI_NAME is declared in the re-exported output module.
    expect(distDts).toMatch(/from '\.\/cli\/output'/);
    const outputDts = fs.readFileSync(
      path.join(path.dirname(distTypes), 'cli', 'output.d.ts'),
      'utf8'
    );
    expect(outputDts).toMatch(/\bCLI_NAME\b/);
  });

  it.runIf(distBuilt)('scratch consumer importing @attentionform/astra typechecks clean', () => {
    expect(() =>
      typecheckConsumer(
        [
          "import { run, CLI_NAME } from '@attentionform/astra';",
          'export function runHelp(): Promise<number> {',
          '  void CLI_NAME;',
          "  return run(['--help']);",
          '}',
          '',
        ].join('\n')
      )
    ).not.toThrow();
  });
});
