import fs from 'node:fs';
import { builtinModules } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, it, expect } from 'vitest';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '../..');

/** Package names a built bundle loads at run time (esbuild leaves them as `require("pkg")`). */
function externalRequires(bundle: string): string[] {
  const source = fs.readFileSync(bundle, 'utf8');
  const names = new Set<string>();
  for (const m of source.matchAll(/\brequire\("([^".][^"]*)"\)/g)) {
    const spec = m[1];
    if (spec.startsWith('node:')) continue;
    const name = spec.startsWith('@') ? spec.split('/').slice(0, 2).join('/') : spec.split('/')[0];
    if (!builtinModules.includes(name)) names.add(name);
  }
  return [...names].sort();
}

// The bundles inline the engine, so what the engine loads is loaded by them too. A package they
// load but do not declare only works where something else happens to have installed it.
describe.each(['cli', 'mcp-server'])('packages/%s declares what its bundle loads', (pkgDir) => {
  const dir = path.join(root, 'packages', pkgDir);
  const bundles = fs.existsSync(path.join(dir, 'dist'))
    ? fs.readdirSync(path.join(dir, 'dist')).filter((f) => f.endsWith('.js')).map((f) => path.join(dir, 'dist', f))
    : [];

  it.skipIf(bundles.length === 0)('every package the bundle requires is a dependency', () => {
    const declared = Object.keys(JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8')).dependencies ?? {});
    const undeclared = bundles.flatMap(externalRequires).filter((name) => !declared.includes(name));
    expect([...new Set(undeclared)]).toEqual([]);
  });
});
